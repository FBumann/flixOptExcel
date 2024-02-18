from typing import Union, List, TypedDict, Optional

import flixOpt.flixComps
import pandas as pd
from typeguard import typechecked
from abc import ABC

from flixOptExcel.Initiation.Modules import ExcelData

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

        self.energy_system: cEnergySystem = self.create_energy_system()

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
        energy_system.addComponents(*list(self.components.values()))
        energy_system.addEffects(*list(self.effects.values()))
        return energy_system

    def create_and_register_components(self):
        for comp_type in self.components_data:
            for component_data in self.components_data[comp_type]:
                if comp_type == "Speicher":     pass
                elif comp_type == "Kessel":     Kessel(component_data).convert_to_model(self.energy_system)
                elif comp_type == "KWK":        pass
                elif comp_type == "KWKekt":     pass
                elif comp_type == "Waermepumpe":pass
                elif comp_type == "EHK":        pass
                elif comp_type == "AbwaermeWP": pass
                elif comp_type == "AbwaermeHT": pass
                elif comp_type == "Rueckkuehler":pass
                else: raise TypeError(f"{comp_type} is not a valid Type of Component. "
                                    f"Implemented types: (KWK, KWKekt, Kessel, EHK, Waermepumpe, "
                                    f"AbwaermeWP, AbwaermeHT, Rueckkuehler, Speicher.")


class DistrictHeatingComponent(ABC):
    @typechecked
    def __init__(self, district_heating_system: DistrictHeatingSystem, time_series_data: pd.DataFrame = None, **kwargs):

        self.district_heating_system: DistrictHeatingSystem = district_heating_system

        self.label: str | None = kwargs.pop('Name')
        self.thermal_power: int | float | None = kwargs.pop('Thermische Leistung')

        self.group: str | None = kwargs.pop('Gruppe', None)

        # Invest
        self.optional: bool | None = kwargs.pop('optional', None)
        self.first_year: int | None = kwargs.pop('first_year', None)
        self.last_year: int | None = kwargs.pop('last_year', None)
        self.costs_fix: int | float | None = kwargs.pop('costs_fix', None)
        self.fund_fix: int | float | None = kwargs.pop('fund_fix', None)
        self.costs_var: int | float | None = kwargs.pop('costs_var', None)
        self.fund_var: int | float | None = kwargs.pop('fund_var', None)
        self.invest_group: cEffectType | None = kwargs.pop('invest_group', None)

        self._kwargs_data: dict = kwargs

    def convert_value_to_TS(self, value: str, time_series_data: pd.DataFrame) -> np.ndarray:
        if value in time_series_data.keys():
            return time_series_data[value].to_numpy()
        else:
            raise Exception(f"{value} is not in given TimeSeries Data.")

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

        '''
        Create an instance of cInvestArgs based on the provided parameters.
        Parameters:
        -----------

        nominal_val : int, float, str, or None
            The nominal value or capacity of the component. If a string is provided, it must be in the format "min-max" to specify a range. If None, investment size is not fixed. optional : bool True if the component allows optional investment, False otherwise. first_year : int The starting year for the component in the calculation. last_year : int The ending year for the component in the calculation.
        costs_fix : float
            Fixed costs associated with the component.
        fund_fix : float
            Fixed funding associated with the component.
        costs_var : float
            Variable costs associated with the component.
        fund_var : float
            Variable funding associated with the component.
        is_flow: bool, optional
            True if the component is a flow, False otherwise. Default is False.
        is_storage : bool, optional
            True if the component is a storage, False otherwise. Default is False.
        is_flow_of_storage : bool, optional
            True if the component is a flow of a storage, False otherwise. Default is False.
        Returns:
        ----------------
        cInvestArgs or (cInvestArgs, cEffectType) if is_flow_of_storage is True
            An instance of cInvestArgs representing the investment parameters for the component. If is_flow_of_storage is True, a tuple
            of (cInvestArgs, cEffectType) is returned.
        Raises:
         ------
        Exception
        - If exactly one of is_flow, is_storage, or is_flow_of_storage is not True.
        - If the format of the nominal_val string is incorrect.
        Notes:
        ---------
            This function creates an instance of cInvestArgs to represent inv
            estment parameters for a component. It calculates the multiplier based on the number of years the component is in the calculation. It also determines if the investment is optional and adjusts costs and funding based on the multiplier. If is_flow_of_storage
            is True, the investment is split between input and output flow components.
        Example usage:
        --------------
            invest_args = get_invest_from_excel(100, False, 2022, 2030, 5000, 10000, 200, 300, is_storage=True) # Returns an instance of cInvestArgs with the specified
            parameters for a storage component.
            invest_args, effect_type = get_invest_from_excel("100-200", True, 2023, 2030, 0, 0, 0, 0, is_flow_of_storage=True) # Returns a tuple containing an instance of cInvestArgs and cEffectType for a flow of storage component.

        '''
        # type checking
        list_of_args = (self.optional, self.costs_fix, self.fund_fix, self.costs_var, self.fund_var)
        if all(value is None for value in list_of_args) and self.nominal_val is not None:
            return None

        # default values
        min_investmentSize = 0
        max_investmentSize = 1e9

        if isinstance(self.nominal_val, (int, float)):
            investmentSize_is_fixed = True
        elif self.nominal_val is None:
            investmentSize_is_fixed = False

        elif isinstance(self.nominal_val, str) and is_valid_format_min_max(self.nominal_val):
            investmentSize_is_fixed = False
            min_investmentSize = float(self.nominal_val.split("-")[0])
            max_investmentSize = float(self.nominal_val.split("-")[1])

        elif isinstance(self.nominal_val, str):
            raise Exception(f"Wrong format of string for nominal_value '{self.nominal_val}'.")
        else:
            raise Exception(f"something went wrong creating the InvestArgs for {self.nominal_val}")

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
                           min_investmentSize=min_investmentSize, max_investmentSize=max_investmentSize)


class Kessel(DistrictHeatingComponent):
    @typechecked
    def __init__(self, **kwargs):
        self.eta_th: float | str = kwargs.pop("eta_th")
        self.fuel_type = kwargs.pop("Brennstoff")
        self.extra_fuel_costs = kwargs.pop("Zusatzkosten pro MWh Brennstoff")
        super().__init__(**kwargs)

    def add_to_system(self, district_heating_system: DistrictHeatingSystem):
        self.exists: int | list[int] | None = self.create_exists(district_heating_system)
        self.investArgs = self.create_invest_args(district_heating_system)
        self.kwargs = self.get_kwargs(district_heating_system)

        return cKessel(label=self.label,
                       group=self.group,
                       eta=self.eta_th,
                       exists=self.exists,
                       Q_th=cFlow(label='Qth',
                                  bus=district_heating_system.busses["Fernwaerme"],
                                  nominal_val=self.nominal_val,
                                  investArgs=self.investArgs,
                                  **self.get_kwargs(district_heating_system)
                                  ),
                       Q_fu=cFlow(label='Qfu',
                                  bus=self.busses[self.fuel_type],
                                  costsPerFlowHour=
                                  self.convert_value_to_TS(self.fuel_type, district_heating_system.timeSeries) +
                                  self.convert_value_to_TS(self.extra_fuel_costs, district_heating_system.timeSeries)
                                  )
                       )


lin = Kessel()
print(lin)
