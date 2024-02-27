# -*- coding: utf-8 -*-
import os
import shutil
from copy import copy
import numpy as np
import pandas as pd
from pprintpp import pprint as pp
from abc import ABC
from typing import Union, List, Any, Type, Tuple, Optional

from flixOpt.flixComps import *
from flixOpt.flixCompsExperimental import KWKektB
from flixOpt.flixStructure import cEffectType, cEnergySystem

from .Modules import ExcelData
from flixOptExcel.Evaluation.flixPostprocessingXL import cModelVisualizer, cVisuData, flixPostXL
from flixOptExcel.Initiation.HelperFcts_in import repeat_elements_of_list, calculate_hourly_rolling_mean, \
    linear_interpolation_with_bounds


class ExcelModel:
    def __init__(self, excel_file_path: str):
        self.excel_data = ExcelData(file_path=excel_file_path)
        self.district_heating_system = DistrictHeatingSystem(self.excel_data)

        self.calc_name = self.excel_data.calc_name
        self.final_directory = os.path.join(self.excel_data.results_directory, self.calc_name)
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
        self._create_dirs_and_save_input_data()

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
        calc_results = self.load_results()

        main_results = calc_results.infos["modboxes"]["info"][0]["main_results"]
        with open(os.path.join(calc_results.folder, "MainResults.txt"), "w") as log_file:
            pp(main_results, log_file)

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

        self.visual_representation.write_html(os.path.join(calc_results.folder, 'Model_structure.html'))

        from flixOptExcel.Evaluation.excelDH import (run_excel_graphics_main,
                                                            run_excel_graphics_years)
        print("START: EXPORT OF RESULTS TO EXCEL...")
        if overview: run_excel_graphics_main(calc_results)
        if annual_results: run_excel_graphics_years(calc_results)

        print("Writing Results to Excel (YE)...")
        if buses_yearly: calc_results.write_bus_results_to_excel("YE")
        if effects_yearly: calc_results.write_effect_results_to_excel( "YE")
        if comps_yearly: calc_results.write_component_results_to_excel( "YE")
        print("...Results to Excel (YE) finished...")

        print("Writing Results to Excel (d)...")
        if buses_daily: calc_results.write_bus_results_to_excel("d")
        if effects_daily: calc_results.write_effect_results_to_excel("d")
        if comps_daily: calc_results.write_component_results_to_excel( "d")
        print("...Results to Excel (d) finished...")

        print("Writing results to Excel (h)...")
        if buses_hourly: calc_results.write_bus_results_to_excel("h")
        if effects_hourly: calc_results.write_effect_results_to_excel( "h")
        if comps_hourly: calc_results.write_component_results_to_excel( "h")
        print("...Results to Excel (h) finished...")

        return calc_results

    def _create_dirs_and_save_input_data(self):
        os.mkdir(self.final_directory)
        input_data_path = os.path.join(self.final_directory, "Inputdata.xlsx")
        shutil.copy2(self.input_excel_file_path, input_data_path)

        with (pd.ExcelWriter(input_data_path, mode="a", engine="openpyxl", if_sheet_exists='overlay') as writer):
            df = self.district_heating_system.time_series_data_internal
            df.to_excel(writer, index=True, sheet_name="Internally_computed_data")

        calc_info = f"""calc = flixPostXL(nameOfCalc='{self.calc_name}', 
        results_folder='{os.path.join(self.final_directory, 'SolveResults')}', 
        outputYears={self.years})"""

        with open(os.path.join(self.final_directory, "calc_info.txt"), "w") as log_file:
            log_file.write(calc_info)

        with open(os.path.join(self.final_directory, "component_data.txt"), "w") as log_file:
            pp(self.excel_data.components_data, log_file)

        with open(os.path.join(self.final_directory, "System_Description.txt"), "w") as log_file:
            print(self.district_heating_system.final_model, file=log_file)

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


class DistrictHeatingSystem:
    def __init__(self, excel_data: ExcelData):
        self.time_series_data = excel_data.time_series_data
        self.time_series_data_internal = pd.DataFrame(index=self.time_series_data.index)
        self.components_data = excel_data.components_data

        self.years = excel_data.years
        self.timeSeries = excel_data.time_series_data.index.to_numpy()
        self.co2_limit_dict = excel_data.co2_limits
        self.co2_factors = excel_data.co2_factors
        self.heating_network_temperature_curves = excel_data.heating_network_temperature_curves

        self._handle_heating_network()

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
                co2_limiter_shares[effects[f"CO2Limit{year}"]] = self.create_exists(first_year=year, last_year=year + 1)
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
        sinks_n_sources['Waermebedarf'] = cSink('Waermebedarf', group="Wärmelast",
                                              sink=cFlow('Qth', bus=self.busses['Fernwaerme'],
                                                         nominal_val=1, val_rel=self.time_series_data["SinkHeat"]))

        sinks_n_sources['Netzverluste'] = cSink('Netzverluste', group="Wärmelast",
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
        pp(self.components_data)
        for comp_type in self.components_data:
            for component_data in self.components_data[comp_type]:
                if comp_type == "Speicher":
                    comp = Storage(**component_data)
                    comp.add_to_model(self)
                elif comp_type == "Kessel":
                    comp = Kessel(**component_data)
                    comp.add_to_model(self)
                elif comp_type == "KWK":
                    comp = KWK(**component_data)
                    comp.add_to_model(self)
                elif comp_type == "KWKekt":
                    comp = KWKekt(**component_data)
                    comp.add_to_model(self)
                elif comp_type == "Waermepumpe":
                    comp = Waermepumpe(**component_data)
                    comp.add_to_model(self)
                elif comp_type == "EHK":
                    comp = EHK(**component_data)
                    comp.add_to_model(self)
                elif comp_type == "AbwaermeWP":
                    comp = AbwaermeWaermepumpe(**component_data)
                    comp.add_to_model(self)
                elif comp_type == "AbwaermeHT":
                    comp = Abwaerme(**component_data)
                    comp.add_to_model(self)
                elif comp_type == "Rueckkuehler":
                    comp = Rueckkuehler(**component_data)
                    comp.add_to_model(self)
                elif comp_type == "Geothermie":
                    comp = Geothermie(**component_data)
                    comp.add_to_model(self)
                else:
                    raise TypeError(f"{comp_type} is not a valid Type of Component. "
                                    f"Implemented types: (KWK, KWKekt, Kessel, EHK, Waermepumpe, "
                                    f"AbwaermeWP, AbwaermeHT, Rueckkuehler, Speicher, Geothermie.")

    def create_exists(self, first_year: Union[int, type(None)], last_year: Union[int, type(None)]) -> np.ndarray:
        '''
        This function was written to match the operation years of a component to the operation years of the system
        ----------

        Returns
        -------
        np.ndarray
        '''
        if first_year is None and last_year is None:
            return np.ones(len(self.time_series_data.index))
        elif first_year is None or last_year is None:
            raise Exception("Either both or none of 'Startjahr' and 'Endjahr' must be set per Component.")
        else:
            # Create a new list with 1s and 0s based on the conditions
            list_to_repeat = [1 if first_year <= num <= last_year else 0 for num in self.years]

            if len(list_to_repeat) == sum(list_to_repeat):
                return np.ones(len(self.time_series_data.index))
            else:
                return np.array(repeat_elements_of_list(list_to_repeat))

    def _handle_heating_network(self):
        """
        # TODO: Redo docstring
        Handle heating network parameters in the input DataFrame.

        This function calculates or checks the presence of key parameters related to the heating network,
        including supply temperature (TVL_FWN), return temperature (TRL_FWN), and network losses (SinkLossHeat).
        If not already present in the dataframe, creates them and returns the filled dataframe


        Raises:
        - Exception: If one of "TVL_FWN" or "TRL_FWN" is not present in the input DataFrame and needs calculation.

        Example:
        ```python
        handle_heating_network(my_dataframe)
        ```

        """

        self.time_series_data['Tamb24mean'] = calculate_hourly_rolling_mean(series=self.time_series_data['Tamb'],
                                                                            window_size=24)
        if "TVL_FWN" in self.time_series_data.keys() and "TRL_FWN" not in self.time_series_data.keys():
            raise Exception("If 'TVL_FWN' is given, 'TRL_FWN' also has to be in the Input Dataset")
        elif "TVL_FWN" not in self.time_series_data.keys() and "TRL_FWN" in self.time_series_data.keys():
            raise Exception("If 'TRL_FWN' is given, 'TVL_FWN' also has to be in the Input Dataset")
        elif "TVL_FWN" and "TRL_FWN" in self.time_series_data.keys():
            print("TVL_FWN and TRL_FWN where included in the input data set")
        else:
            # Berechnung der Vorlauftemperatur
            # TODO: Add Error Handling
            df_tvl = pd.Series()
            i=0
            for year, factors in self.heating_network_temperature_curves.items():
                df = linear_interpolation_with_bounds(input_data=self.time_series_data["Tamb24mean"].iloc[i*8760:(i+1)*8760],
                                                      lower_bound=factors["lb"],
                                                      upper_bound=factors["ub"],
                                                      value_below_bound=factors["value_lb"],
                                                      value_above_bound=factors["value_ub"])
                df_tvl = pd.concat([df_tvl, df])
                i=i+1

            self.time_series_data["TVL_FWN"] = df_tvl

            # TODO: Custom Function?
            self.time_series_data["TRL_FWN"] = np.ones_like(self.time_series_data["TVL_FWN"]) * 60

        if "SinkLossHeat" not in self.time_series_data.keys():  # Berechnung der Netzverluste
            k_loss_netz = 0.4640  # in MWh/K        # Vereinfacht, ohne Berücksichtigung einer sich ändernden Netzlänge
            # TODO: Factor into excel
            self.time_series_data["SinkLossHeat"] = (k_loss_netz *
                                                     ((self.time_series_data["TVL_FWN"] + self.time_series_data["TRL_FWN"]) / 2 -
                                                      self.time_series_data["Tamb"]))
        else:
            print("Heating losses where included in the input data set")


class DistrictHeatingComponent(ABC):
    def __init__(self, **kwargs):
        self._kwargs_data: dict = kwargs

        self.label = self._typechecked_attr('Name', [str])
        self.thermal_power = self._typechecked_attr('Thermische Leistung', [int, float, str], None)
        self.group = self._typechecked_attr('Gruppe', [str], None)

        # Invest
        self.optional = self._typechecked_attr('Optional', [bool], False)

        self.first_year = self._typechecked_attr('Startjahr', [int], None)
        self.lifetime = self._typechecked_attr('Lebensdauer', [int], None)

        self.costs_invest = self._typechecked_attr('Investkosten', [int, float], 0)
        self.costs_invest_specific = self._typechecked_attr('Investkosten/MW', [int, float], 0)
        self.interest_rate = self._typechecked_attr('Zinssatz', [int, float], 0)

        self.costs_operation = self._typechecked_attr('Betriebskosten', [int, float], 0)
        self.cost_operation_specific = self._typechecked_attr('Betriebskosten/MW', [int, float], 0)
        self.costs_grid_fee = self._typechecked_attr('Netzentgelt pro MW', [int, float], 0)

        self.funding_rate = self._typechecked_attr('Fördersatz', [int, float], 0)
        self.invest_group = self._typechecked_attr('Investgruppe', [str], None)

    def kwargs(self, time_series_data:pd.DataFrame) -> dict:
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
                processed_kwargs[key] = self._convert_value_to_TS(value, time_series_data)
            elif value is not None:
                processed_kwargs[key] = value

        if existing_kwargs:  # Check for unprocessed kwargs
            excess_kwargs = ', '.join([f"{key}: {value}" for key, value in existing_kwargs.items()])
            raise Exception(f"Unexpected keyword arguments: {excess_kwargs}")

        return processed_kwargs

    def exists(self, district_heating_system: DistrictHeatingSystem) -> np.ndarray:
        if self.first_year is None and self.lifetime is None:
            return np.ones(len(district_heating_system.time_series_data.index))
        elif self.first_year is None or self.lifetime is None:
            raise Exception("Either both or none of 'Startjahr' and 'Endjahr' must be set per Component.")
        else:
            # Create a new list with 1s and 0s based on the conditions
            list_to_repeat = [1 if self.first_year <= num <= self.first_year + self.lifetime
                              else 0 for num in district_heating_system.years]

            if len(list_to_repeat) == sum(list_to_repeat):
                return np.ones(len(district_heating_system.time_series_data.index))
            else:
                return np.array(repeat_elements_of_list(list_to_repeat))

    def invest_args(self, invest_parameter, district_heating_system: DistrictHeatingSystem) -> Optional[cInvestArgs]:
        if self.first_year is None and self.lifetime is None:
            return None # No invest Args used
        elif self.first_year is None or self.lifetime is None:
            raise ValueError(f"To use Invest Parameters, please set 'Startjahr' and Lebensdauer'")

        invest_values_decoded = self._handle_invest_parameter(invest_parameter)
        investmentSize_is_fixed = isinstance(invest_values_decoded["value"], (int, float))

        fix_costs, specific_costs = self._calculate_costs_and_funding(district_heating_system)

        if self.invest_group is not None:
            specific_costs[district_heating_system.effects[self.invest_group]] = 1

        return cInvestArgs(
            fixCosts={key: value for key, value in fix_costs.items() if value},
            specificCosts={key: value for key, value in specific_costs.items() if value},
            investmentSize_is_fixed=investmentSize_is_fixed,
            investment_is_optional=self.optional,
            min_investmentSize=invest_values_decoded["min"],
            max_investmentSize=invest_values_decoded["max"]
        )

    def _typechecked_attr(self, key, types: List[Type], default: Any = "no_default") -> Any:
        if default is None:
            types.append(type(None))
        elif default == "no_default":
            pass
        else:
            if not isinstance(default, tuple(types)):
                raise TypeError(f"Unexpected default type")

        if default == "no_default":
            try:
                value = self._kwargs_data.pop(key)  # Raises error if key is not present
            except KeyError:
                raise KeyError(
                    f"Key '{key}' not found. Please provide a value for '{key}' of {self.label} of type ({types}).")
        else:
            value = self._kwargs_data.pop(key, default)
        if not isinstance(value, tuple(types)):
            raise TypeError(f"{key} of {self.label} must be one of {types}, not {value} ({type(value)})")
        return value

    def _convert_value_to_TS(self, value: Union[float, int, str], time_series_data: pd.DataFrame) -> np.ndarray:
        if isinstance(value, (int, float)):
            return np.ones(len(time_series_data.index)) * value
        elif value in time_series_data.columns:
            return time_series_data[value].to_numpy()
        else:
            raise Exception(f"{value} of '{self.label}' is not in given TimeSeries Data.")

    def _check_str_format_min_max(self, input_string: str) -> bool:
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

    def _handle_invest_parameter(self, invest_parameter: Union[int, float, str, type(None)]) -> dict:
        '''
        Handles an 'invest_parameter' value by assessing its type and assigning appropriate min, max, and value variables.

        If 'invest_parameter' is string, it should be in the format 'min-max'. If it doesn't follow this format,
        an exception will be raised. In this case, min and max are parsed from the string, and value is set to None.

        If 'invest_parameter' is not string, the min is set to 0, max is set to 1e9 and value is set to the
        'invest_parameter' itself.

        Args:
            invest_parameter (Union[str, int, float]): A number (int, float) or min-max range (string)

        Returns:
            Tuple: Return a tuple containing:
                - value : Value of 'invest_parameter', if it was number. Else, None
                - min   : Minimum limit for 'invest_parameter'
                - max   : Maximum Limit for 'invest_parameter'

        Raises:
            Exception: If 'invest_parameter' is string but does not follow 'min-max' format
        '''
        invest_parameter_decoded = {"value": invest_parameter, "min": 0, "max": 1e9}
        if isinstance(invest_parameter, str):
            if self._check_str_format_min_max(invest_parameter):
                invest_parameter_decoded["min"] = float(invest_parameter.split("-")[0])
                invest_parameter_decoded["max"] = float(self.thermal_power.split("-")[1])
                invest_parameter_decoded["value"] = None
            else:
                raise Exception(f"Wrong format of string for thermal_power '{self.thermal_power}'."
                                f"If thermal power is passed as a string, it must be of the format 'min-max'")

        return invest_parameter_decoded

    def _calculate_costs_and_funding(self, district_heating_system):
        if self.interest_rate == 0: # Preventing ZeroDicvision
            annuity_factor = 1 / self.lifetime
        else:
            annuity_factor = (((1 + self.interest_rate) ** self.lifetime * self.interest_rate) /
                              ((1 + self.interest_rate) ** self.lifetime - 1))
        number_of_years_in_model = sum(1 for num in district_heating_system.years
                                       if self.first_year <= num <= self.first_year + self.lifetime)

        fix_costs = (self.costs_invest * annuity_factor + self.costs_operation) * number_of_years_in_model
        specific_costs = (self.costs_invest_specific * annuity_factor + self.cost_operation_specific +
                          self.costs_grid_fee) * number_of_years_in_model

        fix_costs = {district_heating_system.effects["costs"]: fix_costs,
                     district_heating_system.effects["funding"]: fix_costs * self.funding_rate}
        specific_costs = {district_heating_system.effects["costs"]: specific_costs,
                          district_heating_system.effects["funding"]: specific_costs * self.funding_rate}

        return fix_costs, specific_costs

    def __str__(self):
        """
        Return a readable string representation of the LinearTransformer object,
        showing key attributes.
        """
        attributes = [f"{key}={value}" for key, value in self.__dict__.items()] # TODO: 'if value' to filter out None and 0?
        return f"{self.__class__.__name__}({', '.join(attributes)})"


class KWK(DistrictHeatingComponent):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.eta_th = self._typechecked_attr("eta_th", [float, str])
        self.eta_el = self._typechecked_attr("eta_el", [float, str])
        self.fuel_type = self._typechecked_attr("Brennstoff", [str])
        self.extra_fuel_costs = self._typechecked_attr("Zusatzkosten pro MWh Brennstoff", [int, float], 0)

    def _value_assignment_and_correction(self, time_series_data:pd.DataFrame):
        self.eta_th = self._convert_value_to_TS(self.eta_th, time_series_data)
        self.eta_el = self._convert_value_to_TS(self.eta_el, time_series_data)

        self.costs_grid_fee = self.costs_grid_fee/np.min(self.eta_th + self.eta_el)

    def add_to_model(self, district_heating_system: DistrictHeatingSystem):
        '''
        Adds the Component to the given system
        '''
        self._value_assignment_and_correction(district_heating_system.time_series_data)

        comp = cKWK(label=self.label,
                    group=self.group,
                    eta_th=self.eta_th,
                    eta_el=self.eta_el,
                    exists=self.exists(district_heating_system),
                    Q_th=cFlow(label='Qth',
                               bus=district_heating_system.busses["Fernwaerme"],
                               nominal_val=self._handle_invest_parameter(self.thermal_power)["value"],
                               investArgs=self.invest_args(self.thermal_power, district_heating_system),
                               **self.kwargs(district_heating_system.time_series_data)
                               ),
                    P_el=cFlow(label='Pel',
                               bus=district_heating_system.busses["StromEinspeisung"],
                               costsPerFlowHour=self.electricity_reward_per_flow_hour(district_heating_system)
                               ),
                    Q_fu=cFlow(label='Qfu',
                               bus=district_heating_system.busses[self.fuel_type],
                               costsPerFlowHour=self.fuel_costs_per_flow_hour(district_heating_system)
                               )
                    )

        district_heating_system.final_model.addComponents(comp)

    def electricity_reward_per_flow_hour(self, dhs: DistrictHeatingSystem) -> dict:
        '''
        Information about fuel costs per flow hour
        '''
        return {
            dhs.effects["costs"]: -1 * self._convert_value_to_TS("Strom", dhs.time_series_data),
            dhs.effects["CO2FW"]: self._co2_reward_for_electricity_production(dhs)
        }

    def _co2_reward_for_electricity_production(self, district_heating_system: DistrictHeatingSystem) -> np.ndarray:
        t_vl = district_heating_system.time_series_data["TVL_FWN"].to_numpy() + 273.15
        t_rl = district_heating_system.time_series_data["TRL_FWN"].to_numpy() + 273.15
        t_amb = district_heating_system.time_series_data["Tamb"].to_numpy() + 273.15
        n_el = self.eta_el
        n_th = self.eta_th
        co2_fuel: float = district_heating_system.co2_factors.get("Erdgas", 0)

        # Berechnung der co2-Gutschrift für die Stromproduktion nach der Carnot-Methode
        t_s = (t_vl - t_rl) / np.log((t_vl / t_rl))  # Temperatur Nutzwärme
        n_carnot = 1 - (t_amb / t_s)

        a_el = (1 * n_el) / (n_el + n_carnot * n_th)
        f_el = a_el / n_el
        co2_el = f_el * co2_fuel

        return co2_el

    def fuel_costs_per_flow_hour(self, dhs: DistrictHeatingSystem) -> dict:
        '''
        Information about fuel costs per flow hour
        '''
        costs_dict = {
            dhs.effects["costs"]:
                self._convert_value_to_TS(self.fuel_type, dhs.time_series_data) +
                self._convert_value_to_TS(self.extra_fuel_costs, dhs.time_series_data)
        }

        return costs_dict


class Kessel(DistrictHeatingComponent):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.eta_th = self._typechecked_attr("eta_th", [float, str])
        self.fuel_type = self._typechecked_attr("Brennstoff", [str])
        self.extra_fuel_costs = self._typechecked_attr("Zusatzkosten pro MWh Brennstoff", [int, float], 0)

    def _value_assignment_and_correction(self, time_series_data:pd.DataFrame):
        self.eta_th = self._convert_value_to_TS(self.eta_th, time_series_data)

        self.costs_grid_fee = self.costs_grid_fee/np.min(self.eta_th)

    def fuel_costs_per_flow_hour(self, dhy: DistrictHeatingSystem) -> dict:
        '''
        Information about fuel costs per flow hour
        '''
        costs_dict = {
            dhy.effects["costs"]:
                self._convert_value_to_TS(self.fuel_type, dhy.time_series_data) +
                self._convert_value_to_TS(self.extra_fuel_costs, dhy.time_series_data)
        }

        return costs_dict

    def add_to_model(self, district_heating_system: DistrictHeatingSystem):
        '''
        Adds the Component to the given system
        '''
        self._value_assignment_and_correction(district_heating_system.time_series_data)

        comp = cKessel(label=self.label,
                       group=self.group,
                       eta=self.eta_th,
                       exists=self.exists(district_heating_system),
                       Q_th=cFlow(label='Qth',
                                  bus=district_heating_system.busses["Fernwaerme"],
                                  nominal_val=self._handle_invest_parameter(self.thermal_power)["value"],
                                  investArgs=self.invest_args(self.thermal_power, district_heating_system),
                                  **self.kwargs(district_heating_system.time_series_data)
                                  ),
                       Q_fu=cFlow(label='Qfu',
                                  bus=district_heating_system.busses[self.fuel_type],
                                  costsPerFlowHour=self.fuel_costs_per_flow_hour(district_heating_system)
                                  )
                       )

        district_heating_system.final_model.addComponents(comp)


class Storage(DistrictHeatingComponent):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.capacity = self._typechecked_attr("Kapazität [MWh]", [int, float, str], None)
        self.consider_temperature = self._typechecked_attr("AbhängigkeitVonDT", [bool], False)
        self.losses_per_hour = self._typechecked_attr("VerlustProStunde", [int, float], 0)
        self.eta_load = self._typechecked_attr("eta_load", [int, float], 1)
        self.eta_unload = self._typechecked_attr("eta_unload", [int, float], 1)

        self.costs_capa_invest_specific = self._typechecked_attr("Investkosten/MWh", [int, float], 0)
        self.costs_capa_operation_specific = self._typechecked_attr("Betriebskosten/MWh", [int, float], 0)
        self.invest_group_capa = self._typechecked_attr("Investgruppe Kapazität", [str], None)

    def _value_assignment_and_correction(self, time_series_data:pd.DataFrame):
        self.max_rel = self.calculate_relative_temperature(time_series_data)


    def add_to_model(self, dh_sys: DistrictHeatingSystem):
        '''
        Adds the Component to the given system
        '''
        self._value_assignment_and_correction(dh_sys.time_series_data)

        invest_args_load, invest_args_unload = self.create_invest_args_storage_power(dh_sys)

        storage = cStorage(label=self.label,
                           group=self.group,
                           capacity_inFlowHours=self._handle_invest_parameter(self.capacity)["value"],
                           eta_load=self.eta_load,
                           eta_unload=self.eta_unload,
                           max_rel_chargeState=np.append(self.max_rel, self.max_rel[-1]),
                           exists=self.exists(dh_sys),
                           investArgs=self.invest_args_capacity(dh_sys),
                           inFlow=cFlow(label='QthLoad',
                                        bus=dh_sys.busses["Fernwaerme"],
                                        exists=self.exists(dh_sys),
                                        nominal_val=self._handle_invest_parameter(self.thermal_power)["value"],
                                        max_rel=self.max_rel,
                                        investArgs=invest_args_load,
                                        **self.kwargs(dh_sys.time_series_data)
                                        ),
                           outFlow=cFlow(label='QthUnload',
                                         bus=dh_sys.busses["Fernwaerme"],
                                         exists=self.exists(dh_sys),
                                         nominal_val=self._handle_invest_parameter(self.thermal_power)["value"],
                                         max_rel=self.max_rel,
                                         investArgs=invest_args_unload,
                                         ),
                           avoidInAndOutAtOnce=True,
                           )

        dh_sys.final_model.addComponents(storage)

    def invest_args_capacity(self, district_heating_system: DistrictHeatingSystem) -> Optional[cInvestArgs]:
        if self.first_year is None and self.lifetime is None:
            return None # No invest Args used
        elif self.first_year is None or self.lifetime is None:
            raise ValueError(f"To use Invest Parameters, please set 'Startjahr' and Lebensdauer'")

        invest_values_decoded = self._handle_invest_parameter(self.capacity)
        investmentSize_is_fixed = isinstance(invest_values_decoded["value"], (int, float))

        specific_costs = self._calculate_costs_and_funding_capacity(district_heating_system)

        if self.invest_group_capa is not None:
            specific_costs[district_heating_system.effects[self.invest_group_capa]] = 1

        return cInvestArgs(
            fixCosts={},
            specificCosts={key: value for key, value in specific_costs.items() if value},
            investmentSize_is_fixed=investmentSize_is_fixed,
            investment_is_optional=self.optional,
            min_investmentSize=invest_values_decoded["min"],
            max_investmentSize=invest_values_decoded["max"]
        )

    def create_invest_args_storage_power(self, district_heating_system: DistrictHeatingSystem) -> Tuple[Optional[cInvestArgs], Optional[cInvestArgs]]:
        invest_args = super().invest_args(self.thermal_power, district_heating_system)
        if invest_args is None:
            return None, None
        invest_args_load = copy(invest_args)
        invest_args_unload = copy(invest_args)
        effect_couple_thermal_power = cEffectType(label=f"helpInv{self.label}", unit="",
                                                  description=f"Couple thermal power of in and out flow of {self.label}",
                                                  min_investSum=0, max_investSum=0)
        district_heating_system.final_model.addEffects(effect_couple_thermal_power)

        invest_args_unload.specificCosts = {}  # only include costs in one of the invest costs, else doubled
        invest_args_unload.fixCosts = {}  # only include costs in one of the invest costs, else doubled
        invest_args_unload.specificCosts[effect_couple_thermal_power] = 1
        invest_args_load.specificCosts[effect_couple_thermal_power] = -1

        return invest_args_load, invest_args_unload

    def _calculate_costs_and_funding_capacity(self, district_heating_system):
        if self.interest_rate == 0: # Preventing ZeroDicvision
            annuity_factor = 1 / self.lifetime
        else:
            annuity_factor = (((1 + self.interest_rate) ** self.lifetime * self.interest_rate) /
                              ((1 + self.interest_rate) ** self.lifetime - 1))
        number_of_years_in_model = sum(1 for num in district_heating_system.years
                                       if self.first_year <= num <= self.first_year + self.lifetime)

        specific_costs = (self.costs_capa_invest_specific * annuity_factor +
                          self.costs_capa_operation_specific) * number_of_years_in_model

        specific_costs = {district_heating_system.effects["costs"]: specific_costs,
                          district_heating_system.effects["funding"]: specific_costs * self.funding_rate}

        return specific_costs

    def calculate_relative_temperature(self, time_series_data:pd.DataFrame) -> np.ndarray:
        # TODO: Normalize automatically?
        low_temp = time_series_data["TRL_FWN"].to_numpy()
        high_temp = time_series_data["TVL_FWN"].to_numpy()
        dt_max = 65
        if self.consider_temperature:
            max_rel = ((high_temp - low_temp) / dt_max)
        else:
            max_rel = np.ones(len(time_series_data.index))

        return max_rel


class Waermepumpe(DistrictHeatingComponent):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.cop = self._typechecked_attr("COP", [int, float, str])
        self.calc_cop = self._typechecked_attr("COP berechnen", [bool], False)
        self.operation_funding_bew = self._typechecked_attr("Betriebskostenförderung BEW", [int, float], 0)
        self.lower_bound_of_operation = self._typechecked_attr("Untergrenze für Einsatz", [int, float], None)
        self.lower_bound_ts = self._typechecked_attr("Zeitreihe für Einsatzbeschränkung", [str], None)
        self.extra_fuel_costs = self._typechecked_attr("Zusatzkosten pro MWh Strom", [int, float], 0)

    def _value_assignment_and_correction(self, time_series_data:pd.DataFrame):
        self.cop = self._get_cop_value(time_series_data)
        self.max_rel = self._limit_usage(time_series_data)

        self.costs_grid_fee = self.costs_grid_fee/np.min(self.cop * self.max_rel)

    def add_to_model(self, district_heating_system: DistrictHeatingSystem):
        '''
        Adds the Component to the given system
        '''
        self._value_assignment_and_correction(district_heating_system.time_series_data)

        comp = cHeatPump(label=self.label,
                         group=self.group,
                         COP=self.cop,
                         exists=self.exists(district_heating_system),
                         Q_th=cFlow(label='Qth',
                                    bus=district_heating_system.busses["Fernwaerme"],
                                    nominal_val=self._handle_invest_parameter(self.thermal_power)["value"],
                                    max_rel=self.max_rel,
                                    costsPerFlowHour=self.thermal_reward_per_flow_hour(district_heating_system),
                                    investArgs=self.invest_args(self.thermal_power, district_heating_system),
                                    **self.kwargs(district_heating_system.time_series_data)
                                    ),
                         P_el=cFlow(label='Pel',
                                    bus=district_heating_system.busses["StromBezug"],
                                    costsPerFlowHour=self.electricity_costs_per_flow_hour(district_heating_system)
                                    )
                         )

        district_heating_system.final_model.addComponents(comp)

    def electricity_costs_per_flow_hour(self, dhs: DistrictHeatingSystem) -> dict:
        '''
        Information about fuel costs per flow hour
        '''
        costs_dict = {
            dhs.effects["costs"]:
                self._convert_value_to_TS("Strom", dhs.time_series_data) +
                self._convert_value_to_TS(self.extra_fuel_costs, dhs.time_series_data)
        }

        return costs_dict

    def thermal_reward_per_flow_hour(self, district_heating_system: DistrictHeatingSystem) -> dict:
        '''
        This function calculates the operation funding of a Heat Pump (german program BEW).
        The funding is limited to 90% of electricity costs.

        Parameters
        ----------
        Returns
        -------
        dict
            A dict with the funding per MW_th as a ndarray
        '''

        if self.operation_funding_bew == 0:
            return {}
        else:
            if not isinstance(self.first_year, int):
                raise Exception(f"Need to specify a Year of Operation start for {self.label} to use HP "
                                f"operation funding, because its limited to 10 years.")

            cop = self._get_cop_value(district_heating_system)
            # Förderung pro MW_th
            fund_per_mw_th = (cop - 1 / cop) * self.operation_funding_bew

            # Stromkosten pro MW_th
            costs_for_electricity = self.electricity_costs_per_flow_hour(district_heating_system)[district_heating_system.effects["costs"]]
            el_costs_per_MWth = costs_for_electricity / cop

            # Begrenzung der Förderung auf 90% der Stromkosten
            fund_per_mw_th = np.where(fund_per_mw_th > el_costs_per_MWth * 0.9, el_costs_per_MWth * 0.9,
                                      fund_per_mw_th)

            # Begrenzung auf 10 Jahre
            value_per_year = [1 if self.first_year <= num <= self.first_year + 9 else 0
                              for num in district_heating_system.years]
            correction_factors = np.array(repeat_elements_of_list(value_per_year, 8760))
            fund_per_mw_th = fund_per_mw_th * correction_factors

            return {district_heating_system.effects["funding"]: fund_per_mw_th}

    def _limit_usage(self, time_series_data:pd.DataFrame) -> np.ndarray:
        '''
        Limit useage of a heat pump by a temperature bound

        Returns:
            np.ndarray:
        '''

        if self.lower_bound_of_operation is None:
            return np.ones(len(time_series_data.index))
        else:
            if self.lower_bound_ts is None:
                raise ValueError(f"If a lower bound is given to imit useage of '{self.label}', "
                                 f"you must specify a time series")
            limiting_ts = self._convert_value_to_TS(self.lower_bound_of_operation, time_series_data)
            return np.where(limiting_ts <= self.lower_bound_of_operation, 0, 1)

    def _get_cop_value(self, time_series_data:pd.DataFrame) -> np.ndarray:
        if self.calc_cop:
            if not self.cop in time_series_data.columns:
                raise ValueError(f"The Time Series {self.cop} to calculate the COP of {self.label} could not be found.")
            else:
                return self._calculate_cop(time_series_data[self.cop].to_numpy(),
                                           time_series_data["TVL_FWN"].to_numpy()
                                           )
        else:
            return self._convert_value_to_TS(self.cop, time_series_data)

    def _calculate_cop(self, TqTS:np.ndarray, TsTS:np.ndarray, eta=0.5) -> np.ndarray:
        '''
        Calculates the COP of a heatpump per Timestep from the Temperature of Heat sink and Heat source in Kelvin
        Parameters
        ----------
        TqTS : np.array, float, pd.Dataframe
            Temperature of the Heat Source in Degrees Celcius
        TsTS : np.array, float, pd.Dataframe
            Temperature of the Heat Sink in Degrees Celcius
        eta : float
            Relation to the thermodynamicaly ideal COP

        Returns
        -------
        np.ndarray

        '''
        # Celsius zu Kelvin
        TqTS = TqTS + 273.15
        TsTS = TsTS + 273.15
        COPTS = (TsTS / (TsTS - TqTS)) * eta
        return COPTS


class AbwaermeWaermepumpe(Waermepumpe):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.waste_heat_costs = self._typechecked_attr("Abwärmekosten", [int, float, str])

    def add_to_model(self, district_heating_system: DistrictHeatingSystem):
        '''
        Adds the Component to the given system
        '''
        self._value_assignment_and_correction(district_heating_system.time_series_data)

        comp = cAbwaermeHP(label=self.label,
                           group=self.group,
                           COP=self.cop,
                           exists=self.exists(district_heating_system),
                           Q_th=cFlow(label='Qth',
                                      bus=district_heating_system.busses["Fernwaerme"],
                                      nominal_val=self._handle_invest_parameter(self.thermal_power)["value"],
                                      max_rel=self.max_rel,
                                      investArgs=self.invest_args(self.thermal_power,district_heating_system),
                                      **self.kwargs(district_heating_system.time_series_data)
                                      ),
                           P_el=cFlow(label='Pel',
                                      bus=district_heating_system.busses["StromBezug"],
                                      costsPerFlowHour=self.electricity_costs_per_flow_hour(district_heating_system)
                                      ),
                           Q_ab=cFlow(label='Qab',
                                      bus=district_heating_system.busses["Abwaerme"],
                                      costsPerFlowHour={district_heating_system.effects["costs"]:
                                                            self.waste_heat_costs
                                                        }
                                      )
                           )

        district_heating_system.final_model.addComponents(comp)

class Geothermie(Waermepumpe):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.waste_heat_costs = self._typechecked_attr("Abwärmekosten", [int, float, str])
        self.pump_electricity = self._typechecked_attr("Anteil Pumpstrom", [float])

    def add_to_model(self, district_heating_system: DistrictHeatingSystem):
        '''
        Adds the Component to the given system
        '''
        self._value_assignment_and_correction(district_heating_system.time_series_data)
        comp = cAbwaermeHP(label=self.label,
                           group=self.group,
                           COP=self.cop,
                           exists=self.exists(district_heating_system),
                           Q_th=cFlow(label='Qth',
                                      bus=district_heating_system.busses["Fernwaerme"],
                                      nominal_val=self._handle_invest_parameter(self.thermal_power)["value"],
                                      max_rel=self.max_rel,
                                      investArgs=self.invest_args(self.thermal_power,district_heating_system),
                                      **self.kwargs(district_heating_system.time_series_data)
                                      ),
                           P_el=cFlow(label='Pel',
                                      bus=district_heating_system.busses["StromBezug"],
                                      costsPerFlowHour=self.electricity_costs_per_flow_hour(district_heating_system)
                                      ),
                           Q_ab=cFlow(label='Qab',
                                      bus=district_heating_system.busses["Abwaerme"],
                                      costsPerFlowHour={district_heating_system.effects["costs"]:
                                                            self.waste_heat_costs
                                                        }
                                      )
                           )

        district_heating_system.final_model.addComponents(comp)

    def _get_cop_value(self, time_series_data: pd.DataFrame) -> np.ndarray:
        '''
        Modifies the COP of the heat pump with the amoubt of electricity needed for the geothermal pump
        '''
        cop_hp = super()._get_cop_value(time_series_data)
        return cop_hp/(1 + self.pump_electricity)


class Abwaerme(DistrictHeatingComponent):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.waste_heat_costs = self._typechecked_attr("Abwärmekosten", [int, float, str])

    def add_to_model(self, district_heating_system: DistrictHeatingSystem):
        '''
        Adds the Component to the given system
        '''
        q_th = cFlow(label='Qth',
                     bus=district_heating_system.busses["Fernwaerme"],
                     nominal_val=self._handle_invest_parameter(self.thermal_power)["value"],
                     investArgs=self.invest_args(self.thermal_power,district_heating_system),
                     **self.kwargs(district_heating_system.time_series_data)
                     )

        q_abw = cFlow(label='Qabw',
                      bus=district_heating_system.busses["Abwaerme"],
                      costsPerFlowHour={district_heating_system.effects["costs"]:
                                            self._convert_value_to_TS(self.waste_heat_costs,
                                                                      district_heating_system.time_series_data)
                                        }
                      )

        comp = cBaseLinearTransformer(label=self.label,
                                      group=self.group,
                                      exists=self.exists(district_heating_system),
                                      inputs=[q_abw],
                                      outputs=[q_th],
                                      factor_Sets=[{q_abw: 1, q_th: 1}]
                                      )

        district_heating_system.final_model.addComponents(comp)


class EHK(DistrictHeatingComponent):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.eta_th = self._typechecked_attr("eta_th", [float, str])
        self.extra_fuel_costs = self._typechecked_attr("Zusatzkosten pro MWh Strom", [int, float], 0)

    def fuel_costs_per_flow_hour(self, dhy: DistrictHeatingSystem) -> dict:
        '''
        Information about fuel costs per flow hour
        '''
        costs_dict = {
            dhy.effects["costs"]:
                self._convert_value_to_TS("Strom", dhy.time_series_data) +
                self._convert_value_to_TS(self.extra_fuel_costs, dhy.time_series_data)
        }

        return costs_dict

    def _value_assignment_and_correction(self, time_series_data:pd.DataFrame):
        self.eta_th = self._convert_value_to_TS(self.eta_th, time_series_data)

        self.costs_grid_fee = self.costs_grid_fee/np.min(self.eta_th)

    def add_to_model(self, district_heating_system: DistrictHeatingSystem):
        '''
        Adds the Component to the given system
        '''
        self._value_assignment_and_correction(district_heating_system.time_series_data)

        comp = cEHK(label=self.label,
                    group=self.group,
                    eta=self.eta_th,
                    exists=self.exists(district_heating_system),
                    Q_th=cFlow(label='Qth',
                               bus=district_heating_system.busses["Fernwaerme"],
                               nominal_val=self._handle_invest_parameter(self.thermal_power)["value"],
                               investArgs=self.invest_args(self.thermal_power,district_heating_system),
                               **self.kwargs(district_heating_system.time_series_data)
                               ),
                    P_el=cFlow(label='Pel',
                               bus=district_heating_system.busses["StromBezug"],
                               costsPerFlowHour=self.fuel_costs_per_flow_hour(district_heating_system)
                               )
                    )

        district_heating_system.final_model.addComponents(comp)


class Rueckkuehler(DistrictHeatingComponent):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.electricity_use = self._typechecked_attr("Strombedarf", [int, float], 0)
        self.extra_fuel_costs = self._typechecked_attr("Zusatzkosten pro MWh Strom", [int, float], 0)

    def _value_assignment_and_correction(self, time_series_data:pd.DataFrame):
        self.costs_grid_fee = self.costs_grid_fee * self.electricity_use

    def add_to_model(self, district_heating_system: DistrictHeatingSystem):
        '''
        Adds the Component to the given system
        '''
        self._value_assignment_and_correction(district_heating_system.time_series_data)

        comp = cCoolingTower(label=self.label,
                             group=self.group,
                             exists=self.exists(district_heating_system),
                             specificElectricityDemand=self.electricity_use,
                             Q_th=cFlow(label='Qth',
                                        bus=district_heating_system.busses["Fernwaerme"],
                                        nominal_val=self._handle_invest_parameter(self.thermal_power)["value"],
                                        investArgs=self.invest_args(self.thermal_power,district_heating_system),
                                        **self.kwargs(district_heating_system.time_series_data)
                                        ),
                             P_el=cFlow(label='Pel',
                                        bus=district_heating_system.busses["StromBezug"],
                                        costsPerFlowHour=self.electricity_costs_per_flow_hour(district_heating_system)
                                        )
                             )

        district_heating_system.final_model.addComponents(comp)

    def electricity_costs_per_flow_hour(self, dhs: DistrictHeatingSystem) -> dict:
        '''
        Information about fuel costs per flow hour
        '''
        costs_dict = {
            dhs.effects["costs"]:
                self._convert_value_to_TS("Strom", dhs.time_series_data) +
                self._convert_value_to_TS(self.extra_fuel_costs, dhs.time_series_data)
        }

        return costs_dict


class KWKekt(DistrictHeatingComponent):
    def __init__(self, **kwargs):
        # no super():__init__() needed
        # # Copied from District heating Component. Removed stuff about thermal power.
        self._kwargs_data: dict = kwargs

        self.label = self._typechecked_attr('Name', [str])
        self.group = self._typechecked_attr('Gruppe', [str], None)

        # Invest
        self.optional = self._typechecked_attr('Optional', [bool], None)
        self.first_year = self._typechecked_attr('Startjahr', [int], None)
        self.last_year = self._typechecked_attr('Endjahr', [int], None)
        self.costs_fix = self._typechecked_attr('Fixkosten pro Jahr', [int, float], None)
        self.fund_fix = self._typechecked_attr('Förderung pro Jahr', [int, float], None)
        self.costs_var = self._typechecked_attr('Fixkosten pro MW und Jahr', [int, float], None)
        self.fund_var = self._typechecked_attr('Förderung pro MW und Jahr', [int, float], None)
        self.invest_group = self._typechecked_attr('Investgruppe', [str], None)
        # # Copied from District heating Component. Removed stuff about thermal power.

        self.eta_th = self._typechecked_attr("eta_th", [float, str])
        self.eta_el = self._typechecked_attr("eta_el", [float, str])
        self.fuel_type = self._typechecked_attr("Brennstoff", [str])
        self.extra_fuel_costs = self._typechecked_attr("Zusatzkosten pro MWh Brennstoff", [int, float], 0)

        self.fuel_power = self._typechecked_attr("Brennstoff Leistung", [int, float])
        self.electric_power_points = self._typechecked_attr("Elektrische Leistung (Stützpunkte)", [str])
        self.thermal_power_points = self._typechecked_attr("Thermische Leistung (Stützpunkte)", [str])
        self.can_switch_off = self._typechecked_attr("Ausschaltbar", [bool], True)

        self._validate_power_points()

    def add_to_model(self, district_heating_system: DistrictHeatingSystem):
        '''
        Adds the Component to the given system
        '''

        comp = KWKektB(label=self.label,
                       group=self.group,
                       exists=self.exists(district_heating_system),
                       BusFuel=district_heating_system.busses[self.fuel_type],
                       BusEl=district_heating_system.busses["StromEinspeisung"],
                       BusTh=district_heating_system.busses["Fernwaerme"],
                       nominal_val_Qfu=self.fuel_power,
                       segPel=self.electric_power_points,
                       segQth=self.thermal_power_points,
                       costsPerFlowHour_fuel=self.fuel_costs_per_flow_hour(district_heating_system),
                       costsPerFlowHour_el=self.electricity_reward_per_flow_hour(district_heating_system),
                       iCanSwitchOff=self.can_switch_off,
                       investArgs=self.invest_args(self.fuel_power,district_heating_system),
                       **self.kwargs(district_heating_system.time_series_data)
                       )

        district_heating_system.final_model.addComponents(comp)

    def electricity_reward_per_flow_hour(self, dhs: DistrictHeatingSystem) -> dict:
        '''
        Information about fuel costs per flow hour
        '''
        return {
            dhs.effects["costs"]:
                self._convert_value_to_TS("Strom", dhs.time_series_data) +
                self._convert_value_to_TS(self.extra_fuel_costs, dhs.time_series_data),
        }

    def fuel_costs_per_flow_hour(self, dhs: DistrictHeatingSystem) -> dict:
        '''
        Information about fuel costs per flow hour
        '''
        costs_dict = {
            dhs.effects["costs"]:
                self._convert_value_to_TS(self.fuel_type, dhs.time_series_data) +
                self._convert_value_to_TS(self.extra_fuel_costs, dhs.time_series_data)
        }

        return costs_dict

    def _validate_power_points(self):
        self.electric_power_points = self._string_to_list(self.electric_power_points)
        self.thermal_power_points = self._string_to_list(self.thermal_power_points)
        if len(self.electric_power_points) != len(self.electric_power_points):
            raise ValueError("The number of electricity and thermal power points must be the same")
        for i in range(len(self.electric_power_points)):
            eta_el = self.electric_power_points[i] / self.fuel_power
            eta_th = self.thermal_power_points[i] / self.fuel_power
            if eta_el > 1:
                raise ValueError(f"The electric efficiency of {self.label} exceeds 100%.")
            if eta_th > 1:
                raise ValueError(f"The thermal efficiency of {self.label} exceeds 100%.")

    def _string_to_list(self, delimited_string: str, delimiter: str = '-') -> list:
        """
        Convert a string of hyphen-separated numbers to a list of floats.

        Parameters:
        - delimited_string (str): The input string containing delimited numbers.
        - delimiter (str): The delimiter

        Returns:
        - list: A list of floats representing the numbers in the input string.
        """
        delimited_string = delimited_string.replace(',', '.')
        return list(map(float, delimited_string.split(delimiter)))
