from typing import Union, List, TypedDict, Optional

import os
import shutil

import numpy as np
from pprintpp import pprint as pp
import pandas as pd
from typeguard import typechecked
from abc import ABC

from .Modules import ExcelData

from flixOpt.flixComps import *
from flixOpt.flixStructure import cEffectType, cEnergySystem
from flixOptExcel.Evaluation.HelperFcts_post import flixPostXL
from flixOptExcel.Evaluation.flixPostprocessingXL import cModelVisualizer, cVisuData
from flixOptExcel.Initiation.HelperFcts_in import (check_dataframe_consistency, handle_component_data,
                                                   combine_dicts_of_component_data, convert_component_data_types,
                                                   convert_component_data_for_looping_through,
                                                   calculate_hourly_rolling_mean,
                                                   split_kwargs, create_exists, handle_nom_val, limit_useage,
                                                   calculate_co2_credit_for_el_production,
                                                   is_valid_format_segmentsOfFlows, string_to_list,
                                                   is_valid_format_min_max, createCOPfromTS,
                                                   linear_interpolation_with_bounds, repeat_elements_of_list)


class DistrictHeatingSystem:
    def __init__(self, excel_data: ExcelData):
        self.time_series_data = excel_data.time_series_data
        self.components_data = excel_data.components_data

        self.years = excel_data.years
        self.timeSeries = excel_data.time_series_data.index.to_numpy()
        self.co2_limit_dict = excel_data.co2_limits
        self.co2_factors = excel_data.co2_factors

        self.busses = self.create_busses()

        self.effects = self.create_effects()
        self.sinks_n_sources = self.create_sinks_n_sources()
        self.helpers = self.create_helpers()

        self.final_model: cEnergySystem = self.create_energy_system()

        self.create_and_register_components()

    def create_effects(self) -> dict:
        effects = {}
        effects['target'] = cEffectType('target', 'i.E.', 'Target',  # name, unit, description
                                        isObjective=True)  # defining costs as objective of optimiziation
        effects['costs'] = cEffectType('costs', '€', 'Kosten', isStandard=True,
                                       specificShareToOtherEffects_operation={effects['target']: 1},
                                       specificShareToOtherEffects_invest={effects['target']: 1})

        effects['funding'] = cEffectType('funding', '€', 'Funding Gesamt',
                                         specificShareToOtherEffects_operation={effects['costs']: -1},
                                         specificShareToOtherEffects_invest={effects['costs']: -1})

        effects['CO2FW'] = cEffectType('CO2FW', 't', 'CO2Emissionen der Fernwaerme')

        effects['CO2'] = cEffectType('CO2', 't', 'CO2Emissionen',
                                     specificShareToOtherEffects_operation={
                                         effects['costs']: self.time_series_data["CO2"],
                                         effects['CO2FW']: 1})

        # Limit CO2 Emissions per year
        co2_limiter_shares = {}
        for year in self.years:
            if self.co2_limit_dict.get(year) is not None:
                effects[f"CO2Limit{year}"] = cEffectType(f"CO2Limit{year}", 't',
                                                         description="Effect to limit the Emissions in that year",
                                                         max_operationSum=self.co2_limit_dict[year])
                co2_limiter_shares[effects[f"CO2Limit{year}"]] = create_exists(first_year=year, last_year=year + 1,
                                                                               outputYears=self.years)
        effects['CO2FW'].specificShareToOtherEffects_operation.update(co2_limiter_shares)

        effects.update(self.create_invest_groups())

        return effects

    def create_invest_groups(self):
        effects = {}
        for key, comp_type in self.components_data.items():
            for comp in comp_type:
                label = comp.get("Investgruppe")
                if isinstance(label, str) and label not in effects.keys():
                    limit = float(label.split(":")[-1])
                    label_new = label.replace(":", "")
                    effects[label] = cEffectType(label=label_new, description="Limiting the Investments to 1 per group",
                                                 unit="Stk", max_Sum=limit)
        return effects

    def create_busses(self):
        busses = {}
        excess_costs = None

        busses['StromBezug'] = cBus('el', 'StromBezug', excessCostsPerFlowHour=excess_costs)
        busses['StromEinspeisung'] = cBus('el', 'StromEinspeisung', excessCostsPerFlowHour=excess_costs)
        busses['Fernwaerme'] = cBus('heat', 'Fernwaerme', excessCostsPerFlowHour=excess_costs)
        busses['Erdgas'] = cBus('fuel', 'Erdgas', excessCostsPerFlowHour=excess_costs)
        busses['Wasserstoff'] = cBus('fuel', 'Wasserstoff', excessCostsPerFlowHour=excess_costs)
        busses['EBS'] = cBus(media='fuel', label='EBS', excessCostsPerFlowHour=excess_costs)
        busses['Abwaerme'] = cBus(media='heat', label='Abwaerme', excessCostsPerFlowHour=excess_costs)

        return busses

    def create_sinks_n_sources(self) -> dict:
        sinks_n_sources = {}
        sinks_n_sources['Waermelast'] = cSink('Waermelast', group="Wärmelast_mit_Verlust",
                                              sink=cFlow('Qth', bus=self.busses['Fernwaerme'],
                                                         nominal_val=1, val_rel=self.time_series_data["SinkHeat"]))

        sinks_n_sources['Waermelast_Netzverluste'] = cSink('Waermelast_Netzverluste', group="Wärmelast_mit_Verlust",
                                                           sink=cFlow('Qth', bus=self.busses['Fernwaerme'],
                                                                      nominal_val=1,
                                                                      val_rel=self.time_series_data["SinkLossHeat"]))

        sinks_n_sources['StromSink'] = cSink('StromSink', sink=cFlow('Pel', bus=self.busses['StromEinspeisung']))

        sinks_n_sources['StromSource'] = cSource('StromSource', source=cFlow('Pel', bus=self.busses['StromBezug']))

        sinks_n_sources['ErdgasSource'] = cSource('ErdgasSource',
                                                  source=cFlow('Qfu',
                                                               bus=self.busses['Erdgas'],
                                                               costsPerFlowHour={self.effects["CO2"]:
                                                                                     self.co2_factors.get('Erdgas')}
                                                               )
                                                  )

        sinks_n_sources['WasserstoffSource'] = cSource('WasserstoffSource',
                                                       source=cFlow('Qfu', bus=self.busses['Wasserstoff']))

        sinks_n_sources['EBSSource'] = cSource('EBSSource', source=cFlow('Qfu', bus=self.busses['EBS']))

        sinks_n_sources['AbwaermeSource'] = cSource(label="AbwaermeSource",
                                                    source=cFlow(label='Qabw', bus=self.busses['Abwaerme']))

        return sinks_n_sources

    def create_helpers(self):
        helpers = {}
        Pout1 = cFlow(label="Strompreis",
                      bus=self.busses['StromEinspeisung'],
                      nominal_val=0,
                      costsPerFlowHour=self.time_series_data["Strom"])
        Pout2 = cFlow(label="Gaspreis",
                      bus=self.busses['Erdgas'],
                      nominal_val=0,
                      costsPerFlowHour=self.time_series_data["Erdgas"])
        Pout3 = cFlow(label="Wasserstoffpreis",
                      bus=self.busses['Wasserstoff'],
                      nominal_val=0,
                      costsPerFlowHour=self.time_series_data["Wasserstoff"])
        Pout4 = cFlow(label="EBSPreis",
                      bus=self.busses['EBS'],
                      nominal_val=0,
                      costsPerFlowHour=self.time_series_data["EBS"])
        helpers["HelperPreise"] = cBaseLinearTransformer(label="HelperPreise",
                                                         inputs=[], outputs=[Pout1, Pout2, Pout3, Pout4],
                                                         factor_Sets=[{Pout1: 1, Pout2: 1, Pout3: 1, Pout4: 1}]
                                                         )
        return helpers

    def create_energy_system(self):
        energy_system = cEnergySystem(timeSeries=self.timeSeries)
        energy_system.addEffects(*list(self.effects.values()))
        energy_system.addComponents(*list(self.sinks_n_sources.values()))
        energy_system.addComponents(*list(self.helpers.values()))
        return energy_system

    def create_and_register_components(self):
        for comp_type in self.components_data:
            for component_data in self.components_data[comp_type]:
                if comp_type == "Speicher":
                    comp = Storage(**component_data)
                    comp.add_to_model(self)
                elif comp_type == "Kessel":
                    comp = Kessel(**component_data)
                    comp.add_to_model(self)
                elif comp_type == "KWK":
                    pass
                elif comp_type == "KWKekt":
                    pass
                elif comp_type == "Waermepumpe":
                    pass
                elif comp_type == "EHK":
                    pass
                elif comp_type == "AbwaermeWP":
                    pass
                elif comp_type == "AbwaermeHT":
                    pass
                elif comp_type == "Rueckkuehler":
                    pass
                else:
                    raise TypeError(f"{comp_type} is not a valid Type of Component. "
                                    f"Implemented types: (KWK, KWKekt, Kessel, EHK, Waermepumpe, "
                                    f"AbwaermeWP, AbwaermeHT, Rueckkuehler, Speicher.")


class DistrictHeatingComponent(ABC):
    @typechecked
    def __init__(self, **kwargs):

        self.label: str = kwargs.pop('Name')
        if not isinstance(self.label, str):
            raise TypeError(f"Name of a Component must be a string, not {self.label}({type(self.label)})")
        self.thermal_power: int | float | str | None = kwargs.pop('Thermische Leistung')
        self.group: str | None = kwargs.pop('Gruppe', None)

        # Invest
        self.optional: bool | None = kwargs.pop('Optional', None)
        self.first_year: int | None = kwargs.pop('Startjahr', None)
        self.last_year: int | None = kwargs.pop('Endjahr', None)
        self.costs_fix: int | float | None = kwargs.pop('Fixkosten pro Jahr', None)
        self.fund_fix: int | float | None = kwargs.pop('Förderung pro Jahr', None)
        self.costs_var: int | float | None = kwargs.pop('Fixkosten pro MW und Jahr', None)
        self.fund_var: int | float | None = kwargs.pop('Förderung pro MW und Jahr', None)
        self.invest_group: cEffectType | None = kwargs.pop('Investgruppe', None)

        self._validate_types()

        self._kwargs_data: dict = kwargs

        self.thermal_power_min = 0
        self.thermal_power_max = 1e9
        self.handle_thermal_power()

    def _validate_types(self):
        if not isinstance(self.label, str):
            raise TypeError(f"Name of a Component must be a string, not {self.label}({type(self.label)})")
        if not isinstance(self.thermal_power, (int, float, str, type(None))):
            raise TypeError(f"Name of a Component must be int, float, str or left blank, not {self.thermal_power}({type(self.thermal_power)})")
        if not self.check_str_format_min_max(self.thermal_power):
            raise TypeError(f"If Thermal Power must is a string, it has to be of Format 'min-max' to limit thermal power for investments")
        if not isinstance(self.group, (str, type(None))):
            raise TypeError(f"A Group must be a str or left blank, not {self.group}({type(self.group)})")


    def convert_value_to_TS(self, value: Union[float, int, str], time_series_data: pd.DataFrame) -> Union[
        np.ndarray, float, int]:
        if isinstance(value, (int, float)):
            return value
        if value in time_series_data.keys():
            return time_series_data[value].to_numpy()
        else:
            raise Exception(f"{value} is not in given TimeSeries Data.")

    def check_str_format_min_max(self, input_string: str) -> bool:
        '''
        This function was written to check if a string is of the format "min-max"
        ----------
        Returns
        -------
        bool
        '''
        pattern = r'^\d+-\d+$'
        if re.match(pattern, input_string):
            return True
        else:
            return False

    def handle_thermal_power(self):
        if isinstance(self.thermal_power, (float, int)):
            pass
        elif self.thermal_power is None:
            pass
        elif isinstance(self.thermal_power, str):
            if self.check_str_format_min_max(self.thermal_power):
                self.thermal_power = None
                self.thermal_power_min = float(self.thermal_power.split("-")[0])
                self.thermal_power_max = float(self.thermal_power.split("-")[1])
            else:
                raise Exception(f"Wrong format of string for thermal_power '{self.thermal_power}'."
                                f"If thermal power is passed as a string, it must be of the format 'min-max'")
        else:
            raise Exception(f"Wrong type for invest parameter '{self.thermal_power}:{type(self.thermal_power)}'")

    def to_dict(self) -> dict:
        return self.__dict__

    def __str__(self):
        """
        Return a readable string representation of the LinearTransformer object,
        showing key attributes.
        """
        attributes = [f"{key}={value}" for key, value in self.to_dict().items()]
        return f"{self.__class__.__name__}({', '.join(attributes)})"

    def get_kwargs(self, district_heating_system: DistrictHeatingSystem) -> dict:
        existing_kwargs = self._kwargs_data
        allowed_keys = {
            "min_rel": (int, float, str),
            "max_rel": (int, float, str),
            "costsPerRunningHour": (int, float, str),
            "costsPerFlowHour": (int, float, str),
            "switchOnCosts": (int, float, str),
            "loadFactor_min": (int, float),
            "loadFactor_max": (int, float),
            "onHoursSum_min": (int),
            "onHoursSum_max": (int),
            "onHours_min": (int),
            "onHours_max": (int),
            "offHours_min": (int),
            "offHours_max": (int),
            "switchOn_maxNr": (int),
            "sumFlowHours_min": (int),
            "sumFlowHours_max": (int),
            "iCanSwitchOff": (bool),
        }

        processed_kwargs = {}

        for key, allowed_types in allowed_keys.items():
            value = existing_kwargs.pop(key, None)
            if value is not None and not isinstance(value, allowed_types):
                raise TypeError(f"Expected {key} to be of type {allowed_types}, got {type(value)}")
            elif isinstance(value, str):
                processed_kwargs[key] = self.convert_value_to_TS(value, district_heating_system.time_series_data)
            elif value is not None:
                processed_kwargs[key] = value

        if existing_kwargs:  # Check for unprocessed kwargs
            excess_kwargs = ', '.join([f"{key}: {value}" for key, value in existing_kwargs.items()])
            raise Exception(f"Unexpected keyword arguments: {excess_kwargs}")

        return processed_kwargs

    def create_exists(self, district_heating_system: DistrictHeatingSystem) -> np.ndarray | int | None:
        if self.first_year is None and self.last_year is None:
            return 1
        elif self.first_year is None or self.last_year is None:
            raise Exception("Either both or none of 'Startjahr' and 'Endjahr' must be set per Component.")
        else:
            # Create a new list with 1s and 0s based on the conditions
            list_to_repeat = [1 if self.first_year <= num <= self.last_year else 0 for num in
                              district_heating_system.years]

            if len(list_to_repeat) == sum(list_to_repeat):
                return 1
            else:
                return np.array(repeat_elements_of_list(list_to_repeat))

    def create_invest_args(self, district_heating_system: DistrictHeatingSystem) -> cInvestArgs | None:
        # type checking
        list_of_args = (self.optional, self.costs_fix, self.fund_fix, self.costs_var, self.fund_var)
        if all(value is None for value in list_of_args) and self.thermal_power is not None:
            return None

        if isinstance(self.thermal_power, (int, float)):
            investmentSize_is_fixed = True
        elif self.thermal_power is None:
            investmentSize_is_fixed = False
        else:
            raise Exception(f"something went wrong creating the InvestArgs for {self.label}")

        fixCosts = {district_heating_system.effects["costs"]: self.costs_fix,
                    district_heating_system.effects["funding"]: self.fund_fix}
        specificCosts = {district_heating_system.effects["costs"]: self.costs_var,
                         district_heating_system.effects["funding"]: self.fund_var}

        # Drop if None
        fixCosts = {key: value for key, value in fixCosts.items() if value is not None}
        specificCosts = {key: value for key, value in specificCosts.items() if value is not None}

        # How many years is the comp in the calculation?
        multiplier = sum(
            [1 if self.first_year <= num <= self.last_year else 0 for num in district_heating_system.years])

        # Choose, if it's an optional Investment or a forced investment
        if self.optional:
            investment_is_optional = True
        else:
            investment_is_optional = False

        # Multiply the costs with the number of years the comp is in the calculation
        for key in fixCosts:
            fixCosts[key] *= multiplier
        for key in specificCosts:
            specificCosts[key] *= multiplier

        # Add Investgroup
        if self.invest_group is not None:
            specificCosts[district_heating_system.effects[self.invest_group]] = 1

        return cInvestArgs(fixCosts=fixCosts, specificCosts=specificCosts,
                           investmentSize_is_fixed=investmentSize_is_fixed,
                           investment_is_optional=investment_is_optional,
                           min_investmentSize=self.minimum_thermal_power, max_investmentSize=self.minimum_thermal_power)


class Kessel(DistrictHeatingComponent):
    @typechecked
    def __init__(self, **kwargs):
        self.eta_th: float | str = kwargs.pop("eta_th")
        self.fuel_type = kwargs.pop("Brennstoff")
        self.extra_fuel_costs = kwargs.pop("Zusatzkosten pro MWh Brennstoff")
        super().__init__(**kwargs)

    def add_to_model(self, district_heating_system: DistrictHeatingSystem):
        exists: int | list[int] | None = self.create_exists(district_heating_system)
        invest_args = self.create_invest_args(district_heating_system)
        kwargs = self.get_kwargs(district_heating_system)

        comp = cKessel(label=self.label,
                       group=self.group,
                       eta=self.convert_value_to_TS(self.eta_th, district_heating_system.time_series_data),
                       exists=exists,
                       Q_th=cFlow(label='Qth',
                                  bus=district_heating_system.busses["Fernwaerme"],
                                  nominal_val=self.thermal_power,
                                  investArgs=invest_args,
                                  **kwargs
                                  ),
                       Q_fu=cFlow(label='Qfu',
                                  bus=district_heating_system.busses[self.fuel_type],
                                  costsPerFlowHour=
                                  self.convert_value_to_TS(self.fuel_type, district_heating_system.time_series_data) +
                                  self.convert_value_to_TS(self.extra_fuel_costs,
                                                           district_heating_system.time_series_data)
                                  )
                       )

        district_heating_system.final_model.addComponents(comp)


class KWK(DistrictHeatingComponent):
    @typechecked
    def __init__(self, **kwargs):
        self.eta_th: float | str = kwargs.pop("eta_th")
        self.eta_el: float | str = kwargs.pop("eta_el")
        self.fuel_type = kwargs.pop("Brennstoff")
        self.extra_fuel_costs = kwargs.pop("Zusatzkosten pro MWh Brennstoff")
        super().__init__(**kwargs)

    def add_to_model(self, district_heating_system: DistrictHeatingSystem):
        exists: int | list[int] | None = self.create_exists(district_heating_system)
        invest_args = self.create_invest_args(district_heating_system)
        kwargs = self.get_kwargs(district_heating_system)
        co2_credit = self.co2_credit_for_el(district_heating_system)

        comp = cKWK(label=self.label,
                    group=self.group,
                    eta=self.convert_value_to_TS(self.eta_th, district_heating_system.time_series_data),
                    exists=exists,
                    Q_th=cFlow(label='Qth',
                               bus=district_heating_system.busses["Fernwaerme"],
                               nominal_val=self.thermal_power,
                               investArgs=invest_args,
                               **kwargs
                               ),
                    P_el=cFlow(label='Pel',
                               bus=self.busses["StromEinspeisung"],
                               costsPerFlowHour={self.effects["CO2FW"]: co2_credit,
                                                 self.effects["costs"]: -1 * self.time_series_data["Strom"]}
                               ),
                    Q_fu=cFlow(label='Qfu',
                               bus=district_heating_system.busses[self.fuel_type],
                               costsPerFlowHour=
                               self.convert_value_to_TS(self.fuel_type, district_heating_system.time_series_data) +
                               self.convert_value_to_TS(self.extra_fuel_costs, district_heating_system.time_series_data)
                               )
                    )

        district_heating_system.final_model.addComponents(comp)

    def co2_credit_for_el(self, district_heating_system: DistrictHeatingSystem):
        t_vl = district_heating_system.time_series_data["TVL_FWN"] + 273.15
        t_rl = district_heating_system.time_series_data["TVL_FWN"] + 273.15
        t_amb = district_heating_system.time_series_data["Tamb"] + 273.15
        n_el = self.convert_value_to_TS(self.eta_el, district_heating_system.time_series_data)
        n_th = self.convert_value_to_TS(self.eta_th, district_heating_system.time_series_data)
        co2_fuel: float = district_heating_system.co2_factors.get("Erdgas", 0)

        # Berechnung der co2-Gutschrift für die Stromproduktion nach der Carnot-Methode
        t_s = (t_vl - t_rl) / np.log((t_vl / t_rl))  # Temperatur Nutzwärme
        n_carnot = 1 - (t_amb / t_s)

        a_el = (1 * n_el) / (n_el + n_carnot * n_th)
        f_el = a_el / n_el
        co2_el = f_el * co2_fuel

        return co2_el


class Storage(DistrictHeatingComponent):
    @typechecked
    def __init__(self, **kwargs):
        self.capacity: float | int | None = kwargs.pop("Kapazität [MWh]")
        self.consider_temperature: bool = kwargs.pop("AbhängigkeitVonDT", False)
        self.losses_per_hour: float = kwargs.pop("VerlustProStunde", 0)
        self.eta_load: float = kwargs.pop("eta_load", 1)
        self.eta_unload: float = kwargs.pop("eta_unload", 1)

        self.costs_cap_var: float | None = kwargs.pop('Fixkosten pro MWh und Jahr', None)
        self.fund_cap_var: float | None = kwargs.pop('Förderung pro MWh und Jahr', None)
        super().__init__(**kwargs)

    def create_invest_args_capacity(self, district_heating_system: DistrictHeatingSystem) -> cInvestArgs | None:
        # type checking
        list_of_args = (self.optional, self.costs_cap_var, self.fund_cap_var)
        if all(value is None for value in list_of_args) and self.capacity is not None:
            return None

        # default values
        min_investmentSize = 0
        max_investmentSize = 1e9

        if isinstance(self.capacity, (int, float)):
            investmentSize_is_fixed = True
        elif self.capacity is None:
            investmentSize_is_fixed = False

        elif isinstance(self.capacity, str) and is_valid_format_min_max(self.capacity):
            investmentSize_is_fixed = False
            min_investmentSize = float(self.capacity.split("-")[0])
            max_investmentSize = float(self.capacity.split("-")[1])

        elif isinstance(self.capacity, str):
            raise Exception(f"Wrong format of string for thermal_power '{self.capacity}'.")
        else:
            raise Exception(f"something went wrong creating the InvestArgs for {self.label}")

        fixCosts = None
        specificCosts = {district_heating_system.effects["costs"]: self.costs_cap_var,
                         district_heating_system.effects["funding"]: self.fund_cap_var}

        # Drop if None
        specificCosts = {key: value for key, value in specificCosts.items() if value is not None}

        # How many years is the comp in the calculation?
        multiplier = sum(
            [1 if self.first_year <= num <= self.last_year else 0 for num in district_heating_system.years])

        # Choose, if it's an optional Investment or a forced investment
        if self.optional:
            investment_is_optional = True
        else:
            investment_is_optional = False

        # Multiply the costs with the number of years the comp is in the calculation
        for key in specificCosts:
            specificCosts[key] *= multiplier

        return cInvestArgs(fixCosts=fixCosts, specificCosts=specificCosts,
                           investmentSize_is_fixed=investmentSize_is_fixed,
                           investment_is_optional=investment_is_optional,
                           min_investmentSize=min_investmentSize, max_investmentSize=max_investmentSize)

    def add_to_model(self, district_heating_system: DistrictHeatingSystem):
        exists: int | list[int] | None = self.create_exists(district_heating_system)
        max_rel = self.calculate_relative_capacity_of_storage(
            low_temp=district_heating_system.time_series_data["TRL_FWN"],
            high_temp=district_heating_system.time_series_data["TVL_FWN"],
            dT_max=65)

        # Invest
        invest_args_capacity = self.create_invest_args_capacity(district_heating_system)

        invest_args_load = self.create_invest_args(district_heating_system)
        invest_args_unload = self.create_invest_args(district_heating_system)
        effect_couple_thermal_power=None
        if invest_args_unload is not None:
            effect_couple_thermal_power = cEffectType(label=f"helpInv{self.label}", unit="",
                                                      description=f"Couple thermal power of in and out flow of {self.label}",
                                                      min_investSum=0, max_investSum=0)
            invest_args_load.specificCosts[effect_couple_thermal_power] = -1
            invest_args_unload.specificCosts = {effect_couple_thermal_power: -1}

        kwargs = self.get_kwargs(district_heating_system)

        storage = cStorage(label=self.label,
                           group=self.group,
                           capacity_inFlowHours=self.capacity,
                           eta_load=self.eta_load,
                           eta_unload=self.eta_unload,
                           max_rel_chargeState=max_rel,
                           exists=exists,
                           investArgs=invest_args_capacity,
                           inFlow=cFlow(label='QthLoad',
                                        bus=district_heating_system.busses["Fernwaerme"],
                                        exists=exists,
                                        nominal_val=self.thermal_power,
                                        max_rel=max_rel,
                                        investArgs=invest_args_load,
                                        **kwargs
                                        ),
                           outFlow=cFlow(label='QthUnload',
                                         bus=district_heating_system.busses["Fernwaerme"],
                                         exists=exists,
                                         nominal_val=self.thermal_power,
                                         max_rel=max_rel,
                                         investArgs=invest_args_unload,
                                         ),
                           avoidInAndOutAtOnce=True,
                           )

        if effect_couple_thermal_power is not None:
            district_heating_system.final_model.addEffects(effect_couple_thermal_power)
        district_heating_system.final_model.addComponents(storage)

    def calculate_relative_capacity_of_storage(self, low_temp: np.ndarray, high_temp: np.ndarray,
                                               dT_max: float) -> Union[np.ndarray, float]:
        # TODO: Normalize automatically?
        if self.consider_temperature:
            max_rel = ((high_temp - low_temp) / dT_max)
            return np.append(max_rel, max_rel[-1])
        else:
            return 1


class ExcelModel:
    def __init__(self, excel_file_path: str):
        self.excel_data = ExcelData(file_path=excel_file_path)
        self.district_heating_system = DistrictHeatingSystem(self.excel_data)

        self.calc_name = self.excel_data.calc_name
        self.final_directory: str = os.path.join(self.excel_data.results_directory, self.calc_name)
        self.input_excel_file_path = excel_file_path
        self.years = self.excel_data.years

    @property
    def visual_representation(self):
        visu_data = cVisuData(es=self.district_heating_system.final_model)
        model_visualization = cModelVisualizer(visu_data)
        return model_visualization.Figure

    def print_comps_in_categories(self):
        # String-resources
        print("###############################################")
        print("Initiated Comps:")
        categorized_comps = {}
        for comp in self.district_heating_system.final_model.listOfComponents:
            comp: cBaseComponent
            category = type(comp).__name__
            if category not in categorized_comps:
                categorized_comps[category] = [comp.label]
            else:
                categorized_comps[category].append(comp.label)

        for category, comps in categorized_comps.items():
            print(f"{category}: {comps}")

    def solve_model(self, solver_name: str, gap_frac: float = 0.01, timelimit: int = 3600):
        self.print_comps_in_categories()
        self._adjust_calc_name_and_results_folder()
        self._create_dirs_and_copy_input_excel_file()

        calculation = cCalculation(self.calc_name, self.district_heating_system.final_model, 'pyomo',
                                   pathForSaving=self.final_directory)  # create Calculation
        calculation.doModelingAsOneSegment()

        solver_props = {'gapFrac': gap_frac,  # solver-gap
                        'timelimit': timelimit,  # seconds until solver abort
                        'solver': solver_name,
                        'displaySolverOutput': True,  # ausführlicher Solver-resources.
                        }

        calculation.solve(solver_props, nameSuffix='_' + solver_name,
                          aPath=os.path.join(self.final_directory, "SolveResults"))
        self.calc_name = calculation.nameOfCalc

    def load_results(self) -> flixPostXL:
        return flixPostXL(nameOfCalc=self.calc_name,
                          results_folder=os.path.join(self.final_directory, "SolveResults"),
                          outputYears=self.years)

    def visualize_results(self, overview: bool = True, annual_results: bool = True,
                          buses_yearly: bool = True, comps_yearly: bool = True, effects_yearly: bool = True,
                          buses_daily: bool = True, comps_daily: bool = True, effects_daily: bool = True,
                          buses_hourly: bool = False, comps_hourly: bool = False,
                          effects_hourly: bool = False) -> flixPostXL:
        """
        Visualizes the results of the calculation.

        * The overview results are mainly used to compare yearly mean values
          between different years.

        * The annual results are used to go into detail about the heating
          production and storage usage in each year.

        * The buses results are used to look at all uses of energy balance.

        * The comps results are used to look at all Transformation processes
          in the different components.

        * The effects results are used to look at all effects. Effects are
          Costs, CO2 Funding, etc.

        * Daily mean values are enough for most use cases.

        * Hourly values are good for in-depth examinations, but take a long
          time to extract and save.

        * TAKE CARE: Writing hourly data to excel takes a significant amount of time for
          big Models with many Components.

        Parameters:
            overview (bool): Whether to write overview graphics. Default is True.
            annual_results (bool): Whether to write annual results graphics. Default is True.
            buses_yearly (bool): Whether to write annual results for buses to excel. Default is True.
            comps_yearly (bool): Whether to write annual results for components to excel. Default is True.
            effects_yearly (bool): Whether to write annual results for effects to excel. Default is True.
            buses_daily (bool): Whether to write daily results for buses to excel. Default is True.
            comps_daily (bool): Whether to write daily results for components to excel. Default is True.
            effects_daily (bool): Whether to write daily results for effects to excel. Default is True.
            buses_hourly (bool): Whether to write hourly results for buses to excel. Default is False.
            comps_hourly (bool): Whether to write hourly results for components to excel. Default is False.
            effects_hourly (bool): Whether to write hourly results for effects to excel. Default is False.

        Returns:
            flixPostXL: The calculated results.
        """

        calc_results = self.load_results()

        main_results = calc_results.infos["modboxes"]["info"][0]["main_results"]
        with open(os.path.join(calc_results.folder, "MainResults.txt"), "w") as log_file:
            pp(main_results, log_file)

        self.visual_representation.write_html(os.path.join(calc_results.folder, 'Model_structure.html'))

        from flixOptExcel.Evaluation.graphics_excel import (run_excel_graphics_main,
                                                            run_excel_graphics_years,
                                                            write_bus_results_to_excel,
                                                            write_effect_results_to_excel,
                                                            write_component_results_to_excel)
        print("START: EXPORT OF RESULTS TO EXCEL...")
        if overview: run_excel_graphics_main(calc_results)
        if annual_results: run_excel_graphics_years(calc_results)

        print("Writing Results to Excel (YE)...")
        if buses_yearly: write_bus_results_to_excel(calc_results, "YE")
        if effects_yearly: write_effect_results_to_excel(calc_results, "YE")
        if comps_yearly: write_component_results_to_excel(calc_results, "YE")
        print("...Results to Excel (YE) finished...")

        print("Writing Results to Excel (d)...")
        if buses_daily: write_bus_results_to_excel(calc_results, "d")
        if effects_daily: write_effect_results_to_excel(calc_results, "d")
        if comps_daily: write_component_results_to_excel(calc_results, "d")
        print("...Results to Excel (d) finished...")

        print("Writing results to Excel (h)...")
        if buses_hourly: write_bus_results_to_excel(calc_results, "h")
        if effects_hourly: write_effect_results_to_excel(calc_results, "h")
        if comps_hourly: write_component_results_to_excel(calc_results, "h")
        print("...Results to Excel (h) finished...")

        return calc_results

    def _create_dirs_and_copy_input_excel_file(self):
        os.mkdir(self.final_directory)
        shutil.copy2(self.input_excel_file_path, os.path.join(self.final_directory, "Inputdata.xlsx"))

        calc_info = f"""calc = flixPostXL(nameOfCalc='{self.calc_name}', 
        results_folder='{os.path.join(self.final_directory, 'SolveResults')}', 
        outputYears={self.years})"""

        with open(os.path.join(self.final_directory, "calc_info.txt"), "w") as log_file:
            log_file.write(calc_info)

    def _adjust_calc_name_and_results_folder(self):
        if os.path.exists(self.final_directory):
            for i in range(1, 100):
                calc_name = self.calc_name + "_" + str(i)
                final_directory = os.path.join(os.path.dirname(self.final_directory), calc_name)
                if not os.path.exists(final_directory):
                    self.calc_name = calc_name
                    self.final_directory = final_directory
                    if i >= 5:
                        print(f"There are over {i} different calculations with the same name. "
                              f"Please choose a different name next time.")
                    if i >= 99:
                        raise Exception(f"Maximum number of different calculations with the same name exceeded. "
                                        f"Max is 9999.")
                    break
