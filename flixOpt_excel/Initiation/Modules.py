import os
import shutil
from datetime import datetime, timedelta
from typing import Union, List

import pandas as pd
from pprintpp import pprint as pp

from flixOpt.flixComps import *
from flixOpt.flixStructure import cEffectType, cEnergySystem
from flixOpt_excel.Evaluation.HelperFcts_post import flixPostXL
from flixOpt_excel.Evaluation.flixPostprocessingXL import cModelVisualizer, cVisuData
from .HelperFcts_in import (check_dataframe_consistency, handle_component_data,
                            combine_dicts_of_component_data, convert_component_data_types,
                            convert_component_data_for_looping_through, calculate_hourly_rolling_mean,
                            split_kwargs, create_exists, handle_nom_val, limit_useage,
                            calculate_co2_credit_for_el_production, is_valid_format_segmentsOfFlows, string_to_list,
                            is_valid_format_min_max, createCOPfromTS, linear_interpolation_with_bounds)


class ExcelData:
    def __init__(self, file_path):
        self.file_path = file_path
        self._general_infos = self.read_general_infos()

        # Basic Information
        self.results_directory = self._get_results_directory()
        self.calc_name = self._get_calc_name()

        # Information per year of the Model
        self.years = self._get_years()
        self.co2_limits = self._get_co2_limits()
        self.co2_factors = self._get_co2_factors()

        # Time Series Data
        self.time_series_data = self._read_time_series_data()
        self._further_calculations()

        # Component Data
        self.components_data = self._read_components()


    def read_general_infos(self) -> pd.DataFrame:
        needed_columns =("Jahre", "Zeitreihen Sheets",
                         "Sonstige Zeitreihen Sheets",
                         "Fahrkurve Fernwärmenetz",
                         "CO2-limit",
                         "Erzeuger Sheets",
                         "CO2 Faktor Erdgas [g/MWh_hu]",
                         "Name", "Speicherort")

        general_info = pd.read_excel(self.file_path, sheet_name="Allgemeines")
        general_info = general_info.replace({np.nan: None})

        for column_name in needed_columns:
            if column_name not in general_info:
                raise Exception(f"Column '{column_name}' is missing in sheet 'Allgemeines'.")

        return general_info

    def _get_results_directory(self) -> str:
        path = self._general_infos.at[0, "Speicherort"]

        if not os.path.exists(path):
            raise Exception(f"The path '{path}' for saving does not exist. Please create it first.")
        if not os.path.isdir(path):
            raise Exception(f"The path '{path}' for saving is not a directory.")

        return path

    def _get_calc_name(self) -> str:
        return str(self._general_infos.at[0, "Name"])

    def _get_years(self) -> list:
        years = [year for year in self._general_infos["Jahre"] if year is not None]
        # Type Checking
        for i in range(len(years)):
            if isinstance(years[i], float) and years[i]%int(years[i]) == 0:
                years[i] = int(years[i])
            elif isinstance(years[i], int):
                continue
            else:
                raise ValueError(f"Every year must be an Integer.")
        return years

    def _get_co2_factors(self) -> dict:
        co2_factors = {}
        co2_factors["Erdgas"] = self._general_infos.at[0, "CO2 Faktor Erdgas [g/MWh_hu]"]
        # TODO: Adjust to accept emission factors for multiple Sources
        return co2_factors

    def _get_co2_limits(self) -> dict:
        co2_limit = [limit for limit in self._general_infos["CO2-limit"] if limit is not None]
        # Type Checking
        if not all(isinstance(limit, (int, float, type(None))) for limit in co2_limit):
            raise ValueError(f"Only numbers and Nothing is allowed as CO2-Limit")

        # Checking the number of Limits and filling with None
        missing_limits = len(self.years) - len(co2_limit)
        if missing_limits > 0:
            co2_limit.extend([None] * missing_limits)
        elif missing_limits < 0:
            raise Exception(f"There are more CO2-Limits given than Years where specified.")
        else:
            pass

        return dict(zip(self.years, co2_limit))

    @property
    def _sheetnames_components(self) -> list:
        sheetnames_components = [sheet for sheet in self._general_infos["Erzeuger Sheets"] if sheet is not None]

        if not all(isinstance(name, str) for name in sheetnames_components):
            raise Exception(f"Use Text to specify the Sheetnames of Components")
        if len(sheetnames_components) == 0:
            raise Exception("At least One Sheet Name must be given")

        return sheetnames_components

    @property
    def _sheetnames_ts_data(self) -> list:
        sheetnames_ts_data = [sheet for sheet in self._general_infos["Zeitreihen Sheets"] if sheet is not None]

        if not all(isinstance(name, str) for name in sheetnames_ts_data):
            raise Exception(f"Use Text to specify the Sheetnames of TimeSeries Data")

        # Check if the number of sheetnames matches the number of years
        if not len(sheetnames_ts_data) == len(self.years):
            raise Exception(f"The number of 'years' and the number of 'Zeitreihen Sheets' must match.")

        return sheetnames_ts_data

    @property
    def _sheetnames_ts_data_extra(self) -> list:
        sheetnames_ts_data_extra = [sheet for sheet in self._general_infos["Sonstige Zeitreihen Sheets"] if sheet is not None]

        if not all(isinstance(name, str) for name in sheetnames_ts_data_extra):
            raise Exception(f"Use Text to specify the Sheetnames of TimeSeries Data")

        # Check if the number of sheetnames matches the number of years
        if len(sheetnames_ts_data_extra) == 0:
            pass
        elif len(sheetnames_ts_data_extra) == len(self.years):
            pass
        else:
            raise Exception(f"The number of 'years' and the number of 'Sonstige Zeitreihen Sheets' must match. "
                            f"You can also not use 'Sonstige Zeitreihen Sheets' at all. Just leave the lines blank")

        return sheetnames_ts_data_extra

    @property
    def heating_network_temperature_curves(self) -> dict:
        curve_factors = {}
        curves = [curve for curve in self._general_infos["Fahrkurve Fernwärmenetz"] if curve is not None]

        # Checking the number of curves given
        if len(curves) == 0:
            # TODO: check later, if TVL and T_RL are given!
            return {}

        elif len(curves) == len(self.years):
            if not all(isinstance(curve, str) for curve in curves):
                raise Exception(f"Use Text to specify the Temperature Curve of the heating network. "
                                f"Use Form: ' 'lb'/'value_lb';'ub'/'value_ub' '.")

            pattern = r'^-?\d+/\d+;\d+/\d+$'
            for curve, year in zip(curves, self.years):
                curve_factors[year] = {}
                curve = curve.replace(",", ".")
                curve = curve.replace(" ", "")
                if not re.match(pattern, curve):
                    raise Exception(f"Use Text to specify the Temperature Curve of the heating network. "
                                    f"Use Form: ' 'lb'/'value_lb';'ub'/'value_ub' '."
                                    f"Example:    '-8/120;10/95'.")

                lower, upper = curve.split(";")
                lower_bound, value_below_bound = lower.split("/")
                upper_bound, value_above_bound = upper.split("/")
                curve_factors[year]["lb"] = float(lower_bound)
                curve_factors[year]["ub"] = float(upper_bound)
                curve_factors[year]["value_lb"] = float(value_below_bound)
                curve_factors[year]["value_ub"] = float(value_above_bound)
        else:
            raise Exception(f"The number of temperature curves for the heating network must match the number of years in the Model.")

        return curve_factors

    def _read_time_series_data(self) -> pd.DataFrame:
        li = []  # Initialize an empty list to store DataFrames
        for sheet_name in self._sheetnames_ts_data:
            # Read the Excel sheet, skipping the first two rows, and drop specified columns
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, skiprows=[1, 2]).drop(columns=["Tag", "Uhrzeit"])
            li.append(df)  # Append the DataFrame to the list
        time_series_data = pd.concat(li, axis=0, ignore_index=True)  # Concatenate the DataFrames in the lis

        if len(self._sheetnames_ts_data_extra) > 0:
            li = []  # Initialize an empty list to store DataFrames
            for sheet_name in self._sheetnames_ts_data_extra:
                # Read the Excel sheet, skipping the first two rows, and drop specified columns
                df = pd.read_excel(self.file_path, sheet_name=sheet_name, skiprows=[1, 2]).drop(columns=["Tag", "Uhrzeit"])
                li.append(df)  # Append the DataFrame to the list

            time_series_data_extra = pd.concat(li, axis=0, ignore_index=True)  # Concatenate the DataFrames in the list
            check_dataframe_consistency(df=time_series_data_extra, years=self.years,
                                        name_of_df="time_series_data_extra")

            time_series_data = pd.concat([time_series_data, time_series_data_extra], axis=1)

        check_dataframe_consistency(df=time_series_data, years=self.years)

        # Adding the Index ain datetime format
        a_time_series = datetime(2021, 1, 1) + np.arange(8760*len(self.years)) * timedelta(hours=1)
        a_time_series = a_time_series.astype('datetime64')
        time_series_data.index = a_time_series

        return time_series_data

    def _read_components(self):
        erzeugerdaten = pd.DataFrame()
        for sheet_name in self._sheetnames_components:
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=None, nrows=30)
            erzeugerdaten_new = handle_component_data(df)
            erzeugerdaten = combine_dicts_of_component_data(erzeugerdaten, erzeugerdaten_new)
            print(f"Component Data of Sheet '{sheet_name}' was read sucessfully.")

        erzeugerdaten_converted = convert_component_data_types(erzeugerdaten)
        erz_daten = convert_component_data_for_looping_through(
            erzeugerdaten_converted)  # Keyword Zuweisung und Sortierung
        erz_daten = split_kwargs(erz_daten)
        print("All Component Data was read sucessfully.")
        return erz_daten

    def _further_calculations(self):
        self.time_series_data['Tamb24mean'] = calculate_hourly_rolling_mean(series=self.time_series_data['Tamb'],
                                                                            window_size=24)

        self._handle_heating_network()  # calculation f the heating network temperature and losses
        # TODO rewrite helper fct to return sth instead of overwriting inside the function

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

class ExcelComps:
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
        self.components = self.create_components()

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
                co2_limiter_shares[effects[f"CO2Limit{year}"]] = create_exists({"Startjahr": year, "Endjahr": year},
                                                                               self.years)
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

    def create_components(self):
        finished_components = {}
        for comp_type in self.components_data:
            for component_data in self.components_data[comp_type]:
                if comp_type == "Speicher":
                    comp, bonus_effect = self.create_storage(component_data)
                    if bonus_effect is not None:
                        bonus_effect: cEffectType
                        self.effects[bonus_effect.label] = bonus_effect
                    finished_components[comp.label] = comp
                elif comp_type == "Kessel":
                    comp = self.create_kessel(component_data)
                    finished_components[comp.label] = comp
                elif comp_type == "KWK":
                    comp = self.create_kwk(component_data)
                    finished_components[comp.label] = comp
                elif comp_type == "KWKekt":
                    comps = self.create_kwk_ekt(component_data)
                    for comp in comps:
                        finished_components[comp.label] = comp
                elif comp_type == "Waermepumpe":
                    comp = self.create_heatpump(component_data)
                    finished_components[comp.label] = comp
                elif comp_type == "EHK":
                    comp = self.create_ehk(component_data)
                    finished_components[comp.label] = comp
                elif comp_type == "AbwaermeWP":
                    comp = self.create_abwaermeWP(component_data)
                    finished_components[comp.label] = comp
                elif comp_type == "AbwaermeHT":
                    comp = self.create_abwaermeHT(component_data)
                    finished_components[comp.label] = comp
                elif comp_type == "Rueckkuehler":
                    comp = self.create_rueckkuehler(component_data)
                    finished_components[comp.label] = comp

                else:
                    raise TypeError(f"{comp_type} is not a valid Type of Component. "
                                    f"Implemented types: (KWK, KWKekt, Kessel, EHK, Waermepumpe, "
                                    f"AbwaermeWP, AbwaermeHT, Rueckkuehler, Speicher.")

        finished_components.update(self.create_helpers())

        return finished_components

    # Components #######################################################################################################
    def create_kessel(self, comp_data: dict) -> cKessel:

        invest_args = self.create_investArgs(
            nominal_val=handle_nom_val(comp_data.get("Thermische Leistung")),
            optional=comp_data.get("Optional"),
            startjahr=comp_data.get("Startjahr"), endjahr=comp_data.get("Endjahr"),
            costs_fix=comp_data.get("Fixkosten pro Jahr"),
            fund_fix=comp_data.get("Förderung pro Jahr"),
            costs_var=comp_data.get("Fixkosten pro MW und Jahr"),
            fund_var=comp_data.get("Förderung pro MW und Jahr"),
            invest_group=comp_data.get("Investgruppe")
        )

        kwargs = self.check_and_convert_kwargs(
            kwargs=comp_data["kwargs"],
            existing_keys=["label", "bus", "nominal_val", "investArgs"]
        )

        return cKessel(label=comp_data["Name"],
                       group=comp_data["Gruppe"],
                       eta=self.get_value_or_TS(comp_data["eta_th"]),
                       exists=create_exists(comp_data, self.years),
                       Q_th=cFlow(label='Qth',
                                  bus=self.busses["Fernwaerme"],
                                  nominal_val=handle_nom_val(comp_data.get("Thermische Leistung")),
                                  investArgs=invest_args,
                                  **kwargs
                                  ),
                       Q_fu=cFlow(label='Qfu',
                                  bus=self.busses[comp_data["Brennstoff"]],
                                  costsPerFlowHour=self.get_value_or_TS(comp_data["Brennstoff"]) +
                                                   self.get_value_or_TS(comp_data["Zusatzkosten pro MWh Brennstoff"])
                                  )
                       )

    def create_kwk(self, comp_data: dict) -> Union[cKWK, cBaseLinearTransformer]:

        invest_args = self.create_investArgs(
            nominal_val=handle_nom_val(comp_data.get("Thermische Leistung")),
            optional=comp_data.get("Optional"),
            startjahr=comp_data.get("Startjahr"), endjahr=comp_data.get("Endjahr"),
            costs_fix=comp_data.get("Fixkosten pro Jahr"),
            fund_fix=comp_data.get("Förderung pro Jahr"),
            costs_var=comp_data.get("Fixkosten pro MW und Jahr"),
            fund_var=comp_data.get("Förderung pro MW und Jahr"),
            invest_group=comp_data.get("Investgruppe")
        )

        kwargs = self.check_and_convert_kwargs(
            kwargs=comp_data["kwargs"],
            existing_keys=["label", "bus", "nominal_val", "investArgs"]
        )

        if comp_data["Brennstoff"] == "Erdgas":
            co2_credit = -1 * calculate_co2_credit_for_el_production(
                array_length=len(self.timeSeries),
                t_vl=self.time_series_data["TVL_FWN"],
                t_rl=self.time_series_data["TRL_FWN"],
                t_amb=self.time_series_data["Tamb"],
                n_el=self.get_value_or_TS(comp_data["eta_el"]),
                n_th=self.get_value_or_TS(comp_data["eta_th"]),
                co2_fuel= self.co2_factors.get("Erdgas",0) # TODO: CO2-Faktor von Excel
            )
        else:
            co2_credit = 0

        Q_th = cFlow(label='Qth',
                     bus=self.busses["Fernwaerme"],
                     nominal_val=handle_nom_val(comp_data.get("Thermische Leistung")),
                     investArgs=invest_args,
                     **kwargs
                     )
        Q_fu = cFlow(label='Qfu',
                     bus=self.busses[comp_data["Brennstoff"]],
                     costsPerFlowHour=self.get_value_or_TS(comp_data["Brennstoff"]) +
                                      self.get_value_or_TS(comp_data["Zusatzkosten pro MWh Brennstoff"])
                     )
        P_el = cFlow(label='Pel',
                     bus=self.busses["StromEinspeisung"],
                     costsPerFlowHour={self.effects["CO2FW"]: co2_credit,
                                       self.effects["costs"]: -1 * self.time_series_data["Strom"]}
                     )

        if all(comp_data.get(key) is None for key in ("SegmentsQfu", "SegmentsQth", "SegmentsPel")):  # regular KWK
            return cKWK(label=comp_data["Name"],
                        group=comp_data["Gruppe"],
                        eta_el=self.get_value_or_TS(comp_data["eta_el"]),
                        eta_th=self.get_value_or_TS(comp_data["eta_th"]),
                        exists=create_exists(comp_data, self.years),
                        P_el=P_el, Q_th=Q_th, Q_fu=Q_fu
                        )
        else:
            segQfu = is_valid_format_segmentsOfFlows(comp_data["SegmentsQfu"], mode='decode')
            segQth = is_valid_format_segmentsOfFlows(comp_data["SegmentsQth"], mode='decode')
            segPel = is_valid_format_segmentsOfFlows(comp_data["SegmentsPel"], mode='decode')

            if len(segQfu) == len(segQth) == len(segPel) and len(segQfu) % 2 == 0:
                pass
            else:
                raise Exception("All segments must have the same length and have pairs of values for each segment.")

            segmentsOfFlows = {Q_fu: segQfu, Q_th: segQth, P_el: segPel}

            return cBaseLinearTransformer(
                label=comp_data["Name"],
                group=comp_data["Gruppe"],
                exists=create_exists(comp_data, self.years),
                inputs=[Q_fu],
                outputs=[Q_th, P_el],
                segmentsOfFlows=segmentsOfFlows
            )

    def create_ehk(self, comp_data: dict) -> cEHK:

        invest_args = self.create_investArgs(
            nominal_val=handle_nom_val(comp_data.get("Thermische Leistung")),
            optional=comp_data.get("Optional"),
            startjahr=comp_data.get("Startjahr"), endjahr=comp_data.get("Endjahr"),
            costs_fix=comp_data.get("Fixkosten pro Jahr"),
            fund_fix=comp_data.get("Förderung pro Jahr"),
            costs_var=comp_data.get("Fixkosten pro MW und Jahr"),
            fund_var=comp_data.get("Förderung pro MW und Jahr"),
            invest_group=comp_data.get("Investgruppe")
        )

        kwargs = self.check_and_convert_kwargs(
            kwargs=comp_data["kwargs"],
            existing_keys=["label", "bus", "nominal_val", "investArgs"]
        )

        Q_th = cFlow(label='Qth',
                     bus=self.busses["Fernwaerme"],
                     nominal_val=handle_nom_val(comp_data.get("Thermische Leistung")),
                     investArgs=invest_args,
                     **kwargs
                     )
        P_el = cFlow(label='Pel',
                     bus=self.busses["StromBezug"],
                     costsPerFlowHour={self.effects["costs"]:
                                           self.get_value_or_TS("Strom") +
                                           self.get_value_or_TS(comp_data["Zusatzkosten pro MWh Strom"])
                                       }
                     )

        return cEHK(label=comp_data["Name"],
                    group=comp_data["Gruppe"],
                    eta=self.get_value_or_TS(comp_data["eta_th"]),
                    exists=create_exists(comp_data, self.years),
                    Q_th=Q_th,
                    P_el=P_el,
                    )

    def create_heatpump(self, comp_data: dict) -> cHeatPump:

        existing_keys = ["label", "bus", "nominal_val", "investArgs", "max_rel"]

        invest_args = self.create_investArgs(
            nominal_val=handle_nom_val(comp_data.get("Thermische Leistung")),
            optional=comp_data.get("Optional"),
            startjahr=comp_data.get("Startjahr"), endjahr=comp_data.get("Endjahr"),
            costs_fix=comp_data.get("Fixkosten pro Jahr"),
            fund_fix=comp_data.get("Förderung pro Jahr"),
            costs_var=comp_data.get("Fixkosten pro MW und Jahr"),
            fund_var=comp_data.get("Förderung pro MW und Jahr"),
            invest_group=comp_data.get("Investgruppe")
        )

        funding = self.get_op_fund_bew_of_hp(fund_per_MWamb=comp_data.get("Betriebskostenförderung BEW"),
                                             COP=self.handle_COP_calculation(comp_data["COP"], comp_data["COP berechnen"], comp_data["Name"]),
                                             costs_for_electricity=
                                             self.get_value_or_TS("Strom") +
                                             self.get_value_or_TS(comp_data["Zusatzkosten pro MWh Strom"]),
                                             )

        funding_dict = {}
        if funding is not None:
            existing_keys.append("costsPerFlowHour")
            funding_dict = {self.effects["funding"]: funding}

        kwargs = self.check_and_convert_kwargs(
            kwargs=comp_data["kwargs"],
            existing_keys=existing_keys
        )

        Q_th = cFlow(label='Qth',
                     bus=self.busses["Fernwaerme"],
                     nominal_val=handle_nom_val(comp_data.get("Thermische Leistung")),
                     costsPerFlowHour=funding_dict,
                     max_rel=limit_useage(comp_data, self.time_series_data),
                     investArgs=invest_args,
                     **kwargs
                     )
        P_el = cFlow(label='Pel',
                     bus=self.busses["StromBezug"],
                     costsPerFlowHour={self.effects["costs"]:
                                           self.get_value_or_TS("Strom") +
                                           self.get_value_or_TS(comp_data["Zusatzkosten pro MWh Strom"])
                                       }
                     )

        return cHeatPump(label=comp_data["Name"],
                         group=comp_data["Gruppe"],
                         exists=create_exists(comp_data, self.years),
                         COP=self.handle_COP_calculation(comp_data["COP"], comp_data["COP berechnen"], comp_data["Name"]),
                         Q_th=Q_th,
                         P_el=P_el,
                         )

    def create_kwk_ekt(self, comp_data: dict) -> List[cBaseLinearTransformer]:

        invest_args = self.create_investArgs(
            nominal_val=handle_nom_val(comp_data.get("Brennstoff Leistung")),
            optional=comp_data.get("Optional"),
            startjahr=comp_data.get("Startjahr"), endjahr=comp_data.get("Endjahr"),
            costs_fix=comp_data.get("Fixkosten pro Jahr"),
            fund_fix=comp_data.get("Förderung pro Jahr"),
            costs_var=comp_data.get("Fixkosten pro MW und Jahr"),
            fund_var=comp_data.get("Förderung pro MW und Jahr"),
            invest_group=comp_data.get("Investgruppe")
        )

        kwargs = self.check_and_convert_kwargs(
            kwargs=comp_data["kwargs"],
            existing_keys=["label", "bus", "nominal_val", "investArgs", "iCanSwitchOff", "costPerFlowHour"]
        )

        return KWKektB(label=comp_data["Name"],
                       BusFuel=self.busses[comp_data["Brennstoff"]],
                       BusEl=self.busses["StromEinspeisung"],
                       BusTh=self.busses["Fernwaerme"],
                       exists=create_exists(comp_data, self.years),
                       group=comp_data["Gruppe"],
                       nominal_val_Qfu=comp_data["Brennstoff Leistung"],
                       segPel=string_to_list(comp_data["Elektrische Leistung (Stützpunkte)"]),
                       segQth=string_to_list(comp_data["Thermische Leistung (Stützpunkte)"]),
                       costsPerFlowHour_fuel={self.effects["costs"]:
                                                  self.get_value_or_TS(comp_data["Brennstoff"]) +
                                                  self.get_value_or_TS(comp_data["Zusatzkosten pro MWh Brennstoff"])
                                              },
                       costsPerFlowHour_el={self.effects["costs"]: -1 * self.get_value_or_TS("Strom")},
                       iCanSwitchOff=comp_data.get("canBeTurnedOff", True),
                       investArgs=invest_args,
                       **kwargs
                       )

    def create_abwaermeHT(self, comp_data: dict) -> cBaseLinearTransformer:
        invest_args = self.create_investArgs(
            nominal_val=handle_nom_val(comp_data.get("Thermische Leistung")),
            optional=comp_data.get("Optional"),
            startjahr=comp_data.get("Startjahr"), endjahr=comp_data.get("Endjahr"),
            costs_fix=comp_data.get("Fixkosten pro Jahr"),
            fund_fix=comp_data.get("Förderung pro Jahr"),
            costs_var=comp_data.get("Fixkosten pro MW und Jahr"),
            fund_var=comp_data.get("Förderung pro MW und Jahr"),
            invest_group=comp_data.get("Investgruppe")
        )

        existing_keys = ["label", "bus", "nominal_val", "investArgs"]
        kwargs = self.check_and_convert_kwargs(
            kwargs=comp_data["kwargs"],
            existing_keys=existing_keys
        )

        Q_th = cFlow(label='Qth',
                     bus=self.busses["Fernwaerme"],
                     nominal_val=handle_nom_val(comp_data.get("Thermische Leistung")),
                     investArgs=invest_args,
                     **kwargs
                     )
        Q_abw = cFlow(label='Qabw',
                      bus=self.busses["Abwaerme"],
                      costsPerFlowHour={self.effects["costs"]: self.get_value_or_TS(comp_data["Abwärmekosten"])}
                      )

        return cBaseLinearTransformer(label=comp_data["Name"],
                                      group=comp_data["Gruppe"],
                                      exists=create_exists(comp_data, self.years),
                                      inputs=[Q_abw], outputs=[Q_th], factor_Sets=[{Q_abw: 1, Q_th: 1}]
                                      )

    def create_abwaermeWP(self, comp_data: dict) -> cBaseLinearTransformer:
        invest_args = self.create_investArgs(
            nominal_val=handle_nom_val(comp_data.get("Thermische Leistung")),
            optional=comp_data.get("Optional"),
            startjahr=comp_data.get("Startjahr"), endjahr=comp_data.get("Endjahr"),
            costs_fix=comp_data.get("Fixkosten pro Jahr"),
            fund_fix=comp_data.get("Förderung pro Jahr"),
            costs_var=comp_data.get("Fixkosten pro MW und Jahr"),
            fund_var=comp_data.get("Förderung pro MW und Jahr"),
            invest_group=comp_data.get("Investgruppe")
        )

        funding = self.get_op_fund_bew_of_hp(fund_per_MWamb=comp_data.get("Betriebskostenförderung BEW"),
                                             COP=self.handle_COP_calculation(comp_data["COP"], comp_data["COP berechnen"], comp_data["Name"]),
                                             costs_for_electricity=
                                             self.get_value_or_TS("Strom") +
                                             self.get_value_or_TS(comp_data["Zusatzkosten pro MWh Strom"])
                                             )

        existing_keys = ["label", "bus", "nominal_val", "investArgs", "max_rel"]
        if funding is not None:
            existing_keys.append("costsPerFlowHour")

        kwargs = self.check_and_convert_kwargs(
            kwargs=comp_data["kwargs"],
            existing_keys=existing_keys
        )

        Q_th = cFlow(label='Qth',
                     bus=self.busses["Fernwaerme"],
                     nominal_val=handle_nom_val(comp_data.get("Thermische Leistung")),
                     costsPerFlowHour={self.effects["funding"]: funding},
                     max_rel=limit_useage(comp_data, self.time_series_data),
                     investArgs=invest_args,
                     **kwargs
                     )
        P_el = cFlow(label='Pel',
                     bus=self.busses["StromBezug"],
                     costsPerFlowHour={self.effects["costs"]:
                                           self.get_value_or_TS("Strom") +
                                           self.get_value_or_TS(comp_data["Zusatzkosten pro MWh Strom"])
                                       }
                     )
        Q_abw = cFlow(label='Qabw',
                      bus=self.busses["Abwaerme"],
                      costsPerFlowHour={
                          self.effects["costs"]: self.get_value_or_TS(comp_data["Abwärmekosten"])}
                      )

        return cAbwaermeHP(label=comp_data["Name"],
                           group=comp_data["Gruppe"],
                           exists=create_exists(comp_data, self.years),
                           COP=self.handle_COP_calculation(comp_data["COP"], comp_data["COP berechnen"], comp_data["Name"]),
                           Q_th=Q_th,
                           P_el=P_el,
                           Q_ab=Q_abw,
                           )

    def create_rueckkuehler(self, comp_data: dict) -> cCoolingTower:

        invest_args = self.create_investArgs(
            nominal_val=handle_nom_val(comp_data.get("Thermische Leistung")),
            optional=comp_data.get("Optional"),
            startjahr=comp_data.get("Startjahr"), endjahr=comp_data.get("Endjahr"),
            costs_fix=comp_data.get("Fixkosten pro Jahr"),
            fund_fix=comp_data.get("Förderung pro Jahr"),
            costs_var=comp_data.get("Fixkosten pro MW und Jahr"),
            fund_var=comp_data.get("Förderung pro MW und Jahr"),
            invest_group=comp_data.get("Investgruppe")
        )

        existing_keys = ["label", "bus", "nominal_val", "investArgs", "max_rel"]

        kwargs = self.check_and_convert_kwargs(
            kwargs=comp_data["kwargs"],
            existing_keys=existing_keys
        )

        # Beschränkung Tageszeit: Nur Stunde 8-20
        if comp_data['Beschränkung Einsatzzeit']:
            max_rel = np.tile(np.concatenate((np.zeros(8), np.ones(12), np.zeros(4))), int(len(self.timeSeries) / 24))
        else:
            max_rel = 1

        Q_th = cFlow(label='Qth',
                     bus=self.busses["Fernwaerme"],
                     nominal_val=handle_nom_val(comp_data.get("Thermische Leistung")),
                     max_rel=max_rel,
                     costsPerRunningHour=comp_data["KostenProBetriebsstunde"],
                     investArgs=invest_args,
                     **kwargs
                     )
        P_el = cFlow(label='Pel',
                     bus=self.busses["StromBezug"],
                     costsPerFlowHour={self.effects["costs"]:
                                           self.get_value_or_TS("Strom") +
                                           self.get_value_or_TS(comp_data["Zusatzkosten pro MWh Strom"])
                                       }
                     )

        return cCoolingTower(label=comp_data["Name"],
                             group=comp_data["Gruppe"],
                             exists=create_exists(comp_data, self.years),
                             specificElectricityDemand=comp_data.get("Strombedarf"),
                             Q_th=Q_th,
                             P_el=P_el
                             )

    def create_storage(self, comp_data: dict):

        invest_args_capacity = self.create_investArgs(
            nominal_val=handle_nom_val(comp_data.get("Kapazität [MWh]")),
            optional=comp_data.get("Optional"),
            startjahr=comp_data.get("Startjahr"), endjahr=comp_data.get("Endjahr"),
            costs_fix=comp_data.get("Fixkosten pro Jahr"),
            fund_fix=comp_data.get("Förderung pro Jahr"),
            costs_var=comp_data.get("Fixkosten pro MWh und Jahr"),
            fund_var=comp_data.get("Förderung pro MWh und Jahr"),
            invest_group=comp_data.get("Investgruppe")
        )

        invest_args_flow_in = self.create_investArgs(
            nominal_val=handle_nom_val(comp_data.get("Lade/Entladeleistung [MW]")),
            optional=comp_data.get("Optional"),
            startjahr=comp_data.get("Startjahr"), endjahr=comp_data.get("Endjahr"),
            costs_fix=None,
            fund_fix=None,
            costs_var=comp_data.get("Fixkosten pro MW und Jahr"),
            fund_var=comp_data.get("Förderung pro MW und Jahr"),
        )

        # invest_args_flow_out = copy(invest_args_flow_in) # TODO: is this safe?
        invest_args_flow_out = self.create_investArgs(
            nominal_val=handle_nom_val(comp_data.get("Lade/Entladeleistung [MW]")),
            optional=comp_data.get("Optional"),
            startjahr=comp_data.get("Startjahr"), endjahr=comp_data.get("Endjahr"),
            costs_fix=None,
            fund_fix=None,
            costs_var=comp_data.get("Fixkosten pro MW und Jahr"),
            fund_var=comp_data.get("Förderung pro MW und Jahr"),
        )
        if invest_args_flow_in is None and invest_args_flow_out is None:
            Invest_effect = None
        else:
            # Couple the in and out Flow
            Invest_effect = cEffectType(label=f"helpInv{comp_data['Name']}", unit="", description="InvestHelp",
                                        min_investSum=0, max_investSum=0)
            invest_args_flow_in.specificCosts[Invest_effect] = -1
            invest_args_flow_out.specificCosts[Invest_effect] = 1

        capacity_max_rel = self.calculate_relative_capacity_of_storage(comp_data["AbhängigkeitVonDT"], 65)
        max_rel_flows = capacity_max_rel[:-1]

        existing_keys = ["label", "bus", "exists", "nominal_val", "investArgs", "max_rel"]

        kwargs = self.check_and_convert_kwargs(
            kwargs=comp_data["kwargs"],
            existing_keys=existing_keys
        )

        Q_in = cFlow(label="QthLoad",
                     bus=self.busses["Fernwaerme"],
                     exists=create_exists(comp_data, self.years),
                     nominal_val=handle_nom_val(comp_data["Lade/Entladeleistung [MW]"]),
                     max_rel=max_rel_flows,
                     investArgs=invest_args_flow_in,
                     **kwargs
                     )

        Q_out = cFlow(label="QthUnload",
                      bus=self.busses["Fernwaerme"],
                      max_rel=max_rel_flows,
                      nominal_val=handle_nom_val(comp_data["Lade/Entladeleistung [MW]"]),
                      exists=create_exists(comp_data, self.years),
                      investArgs=invest_args_flow_out,
                      **kwargs
                      )
        storage = cStorage(label=comp_data["Name"],
                           group=comp_data["Gruppe"],
                           exists=create_exists(comp_data, self.years),
                           eta_load=comp_data.get("eta_load", 1),
                           eta_unload=comp_data.get("eta_unload", 1),
                           inFlow=Q_in,
                           outFlow=Q_out,
                           avoidInAndOutAtOnce=True,
                           max_rel_chargeState=capacity_max_rel, fracLossPerHour=comp_data["VerlustProStunde"],
                           capacity_inFlowHours=handle_nom_val(comp_data["Kapazität [MWh]"]),
                           investArgs=invest_args_capacity
                           )

        return [storage, Invest_effect]

    # Utilities ########################################################################################################
    def get_value_or_TS(self, value: Union[str, int, float]) -> Union[np.ndarray, int, float]:
        if isinstance(value, (int, float, bool)):
            return value
        elif isinstance(value, str):
            if value in self.time_series_data.keys():
                return self.time_series_data[value].to_numpy()
            else:
                raise Exception(f"{value} is not in given TimeSeries Data.")
        else:
            raise Exception(f"{type(value)} is not a valid Type. Must be string, number or bool")

    def check_and_convert_kwargs(self, kwargs: dict, existing_keys: list) -> dict:
        '''

        Parameters
        ----------
        kwargs: the kwargs, stored in a dict.
        existing_keys: Kwargs that are not allowed

        Returns
        -------

        '''
        new_kwargs = {}
        possible_kwargs = ("min_rel",
                           "max_rel",
                           "loadFactor_min",
                           "loadFactor_max",
                           "onHoursSum_min",
                           "onHoursSum_max",
                           "onHours_min",
                           "onHours_max",
                           "offHours_min",
                           "offHours_max",
                           "switchOn_maxNr",
                           "sumFlowHours_min",
                           "sumFlowHours_max",
                           "costsPerRunningHour",
                           "costsPerFlowHour",
                           "switchOnCosts",
                           "iCanSwitchOff")

        # Convert the kwargs to TS, if a string is given
        for key, value in kwargs.items():
            if key in possible_kwargs:
                new_kwargs[key] = self.get_value_or_TS(value)
            else:
                raise Exception(f"Keyword Argument {key} not allowed. Choose from {str(possible_kwargs)}.")

        for key, value in new_kwargs.items():
            if key in existing_keys:
                raise Exception(f"Keyword Argument {key} is already used through other Arguments.")

        return new_kwargs

    def create_investArgs(self, nominal_val, optional: bool, startjahr: int, endjahr: int,
                          costs_fix: float, fund_fix: float, costs_var: float, fund_var: float,
                          invest_group: cEffectType = None,
                          is_flow_of_storage: bool = False) -> Union[cInvestArgs, None]:

        '''
        Create an instance of cInvestArgs based on the provided parameters.
        Parameters:
        -----------

        nominal_val : int, float, str, or None
            The nominal value or capacity of the component. If a string is provided, it must be in the format "min-max" to specify a range. If None, investment size is not fixed. optional : bool True if the component allows optional investment, False otherwise. startjahr : int The starting year for the component in the calculation. endjahr : int The ending year for the component in the calculation.
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
        list_of_args = (optional, costs_fix, fund_fix, costs_var, fund_var)
        if all(value is None for value in list_of_args) and nominal_val is not None:
            return None

        # default values
        min_investmentSize = 0
        max_investmentSize = 1e9

        if isinstance(nominal_val, (int, float)):
            investmentSize_is_fixed = True
        elif nominal_val is None:
            investmentSize_is_fixed = False

        elif isinstance(nominal_val, str) and is_valid_format_min_max(nominal_val):
            investmentSize_is_fixed = False
            min_investmentSize = float(nominal_val.split("-")[0])
            max_investmentSize = float(nominal_val.split("-")[1])

        elif isinstance(nominal_val, str):
            raise Exception(f"Wrong format of string for nominal_value '{nominal_val}'.")
        else:
            raise Exception(f"something went wrong creating the InvestArgs for {nominal_val}")

        fixCosts = {self.effects["costs"]: costs_fix,
                    self.effects["funding"]: fund_fix}
        specificCosts = {self.effects["costs"]: costs_var,
                         self.effects["funding"]: fund_var}

        # Drop if None
        fixCosts = {key: value for key, value in fixCosts.items() if value is not None}
        specificCosts = {key: value for key, value in specificCosts.items() if value is not None}

        # How many years is the comp in the calculation?
        multiplier = sum([1 if startjahr <= num <= endjahr else 0 for num in self.years])
        # Fallunterschiedung
        if is_flow_of_storage:
            multiplier = multiplier * 0.5  # because investment is split between the input and output flow

        # Choose, if it's an optional Investment or a forced investment
        if optional:
            investment_is_optional = True
        else:
            investment_is_optional = False

        # Multiply the costs with the number of years the comp is in the calculation
        for key in fixCosts:
            fixCosts[key] *= multiplier
        for key in specificCosts:
            specificCosts[key] *= multiplier

        # Add Investgroup
        if invest_group is not None:
            specificCosts[self.effects[invest_group]] = 1

        Invest = cInvestArgs(fixCosts=fixCosts, specificCosts=specificCosts,
                             investmentSize_is_fixed=investmentSize_is_fixed,
                             investment_is_optional=investment_is_optional,
                             min_investmentSize=min_investmentSize, max_investmentSize=max_investmentSize)
        return Invest

    def get_op_fund_bew_of_hp(self, COP: Union[np.ndarray, cTSraw, float], costs_for_electricity: np.ndarray,
                              fund_per_MWamb: float = None) -> Union[np.ndarray, None]:
        '''
        This function was written to calculate the operation funding of a Heat Pump (BEW)

        ----------

        Returns
        -------
        dict
        '''
        # Betriebskostenförderung, Beschränkt auf 90% der Stromkosten
        if fund_per_MWamb is None:
            return None
        else:
            if isinstance(COP, cTSraw):
                COP = COP.value
            fund_per_MWamb = self.get_value_or_TS(fund_per_MWamb)
            # Förderung pro MW_th
            fund_per_mwth = (COP - 1 / COP) * fund_per_MWamb

            # Stromkosten pro MW_th
            el_costs_per_MWth = costs_for_electricity / COP

            # Begrenzung der Förderung auf 90% der Stromkosten
            return np.where(fund_per_mwth > el_costs_per_MWth * 0.9, el_costs_per_MWth * 0.9, fund_per_mwth)

    def calculate_relative_capacity_of_storage(self, calculate_DT: bool, dT_max: float = 65) -> Union[list, int]:
        '''
        This function was written to calculate the relative capacity of a Storage due to the changing
        temperature Spread in a Heating network
        ----------

        Returns
        -------
        list
        '''
        if calculate_DT:
            maxReldT = ((self.time_series_data["TVL_FWN"] - self.time_series_data["TRL_FWN"]) / dT_max).values.tolist()
            maxReldT.append(maxReldT[-1])
        else:
            maxReldT = 1

        return maxReldT

    def handle_COP_calculation(self, COP: Union[int, float, str], calc_COP_from_TS: bool, name_of_comp:str, eta_carnot=0.5) -> Union[cTSraw, float]:
        '''
        This function was written to assign a COP to a Heat Pump
        ----------

        Returns
        -------
        (fuel_bus: cBus, fuel_costs: dict)
        '''
        # Wenn fixer COP übergeben wird
        if isinstance(COP, (int, float)):
            return COP
        elif isinstance(COP, str) and COP in self.time_series_data.keys():  # Wenn verlinkung zu Temperatur der waermequelle vorgegeben ist
            if calc_COP_from_TS:
                COP = createCOPfromTS(TqTS=self.time_series_data[COP], TsTS=self.time_series_data["TVL_FWN"], eta=eta_carnot)
            else:
                COP = self.time_series_data[COP]
            self.time_series_data["COP" + name_of_comp] = COP
        else:
            raise Exception("Verlinkung zwischen COP der WP " + name_of_comp + " und der Zeitreihe ist fehlgeschlagen. Prüfe den Namen der Zeitreihe")

        return cTSraw(COP)

class EnergySystemExcel(cEnergySystem):
    def __init__(self, excel_comps:ExcelComps):
        super().__init__(timeSeries=excel_comps.timeSeries)
        self.addEffects(*list(excel_comps.effects.values()))
        self.addComponents(*list(excel_comps.components.values()))
        self.addComponents(*list(excel_comps.sinks_n_sources.values()))

    def addEffects(self, *args):
        list_of_effects = list(args)
        #TODO: logic to rorder the effects, so no effect gets added before its shares are already added
        super().addEffects(*list_of_effects)

class ExcelModel:
    def __init__(self, excel_file_path: str):
        self.excel_data = ExcelData(file_path=excel_file_path)
        self.excel_comps = ExcelComps(excel_data=ExcelData(file_path=excel_file_path))

        self.energy_system = EnergySystemExcel(excel_comps=self.excel_comps)

        self.calc_name = self.excel_data.calc_name
        self.final_directory = os.path.join(self.excel_data.results_directory, self.calc_name)
        self.input_excel_file_path = excel_file_path
        self.years = self.excel_data.years

        self.visual_representation.show()



    @property
    def visual_representation(self):
        visu_data = cVisuData(es=self.energy_system)
        model_visualization = cModelVisualizer(visu_data)
        return model_visualization.Figure

    def print_comps_in_categories(self):
        # String-resources
        print("###############################################")
        print("Initiated Comps:")
        categorized_comps = {}
        for comp in self.energy_system.listOfComponents:
            comp: cBaseComponent
            category = type(comp).__name__
            if category not in categorized_comps:
                categorized_comps[category] = [comp.label]
            else:
                categorized_comps[category].append(comp.label)

        for category, comps in categorized_comps.items():
            print(f"{category}: {comps}")

    def solve_model(self, solver_name:str, gap_frac:float=0.01, timelimit:int= 3600):
        self.print_comps_in_categories()
        self._adjust_calc_name_and_results_folder()
        self._create_dirs_and_copy_input_excel_file()

        calculation = cCalculation(self.calc_name, self.energy_system, 'pyomo',
                                   pathForSaving=self.final_directory)  # create Calculation
        calculation.doModelingAsOneSegment()

        solver_props = {'gapFrac': gap_frac,  # solver-gap
                        'timelimit': timelimit,  # seconds until solver abort
                        'solver': solver_name,
                        'displaySolverOutput': True,  # ausführlicher Solver-resources.
                        }

        calculation.solve(solver_props, nameSuffix='_' + solver_name, aPath=os.path.join(self.final_directory, "SolveResults"))
        self.calc_name = calculation.nameOfCalc

    def load_results(self) -> flixPostXL:
        return flixPostXL(nameOfCalc=self.calc_name,
                          results_folder=os.path.join(self.final_directory, "SolveResults"),
                          outputYears=self.years)
    def visualize_results(self) ->flixPostXL:
        calc_results = self.load_results()

        main_results = calc_results.infos["modboxes"]["info"][0]["main_results"]
        with open(os.path.join(calc_results.folder, "MainResults.txt"), "w") as log_file:
            pp(main_results, log_file)

        self.visual_representation.write_html(os.path.join(calc_results.folder, 'Model_structure.html'))

        from flixOpt_excel.Evaluation.graphics_excel import (run_excel_graphics_main,
                                                             run_excel_graphics_years,
                                                             save_in_n_outputs_per_comp_and_bus_and_effects)
        run_excel_graphics_main(calc_results)
        run_excel_graphics_years(calc_results)
        save_in_n_outputs_per_comp_and_bus_and_effects(calc_results, buses=True, comps=True, effects=True, resample_by="D")

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
            for i in range(1,100):
                calc_name = self.calc_name + "_" + str(i)
                final_directory = os.path.join(os.path.dirname(self.final_directory), calc_name)
                if not os.path.exists(final_directory):
                    self.calc_name = calc_name
                    self.final_directory = final_directory
                    if i >= 5:
                        print(f"There are over {i} different calculations with the same name. "
                              f"Please choose a different name next time.")
                    if i>= 99:
                        raise Exception(f"Maximum number of different calculations with the same name exceeded. "
                                        f"Max is 9999.")
                    break