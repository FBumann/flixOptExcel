import os
import shutil
from datetime import datetime, timedelta
from typing import Union, List

import pandas as pd
from pprintpp import pprint as pp

from flixOpt.flixComps import *
from flixOpt.flixStructure import cEffectType, cEnergySystem
from flixOptExcel.Evaluation.HelperFcts_post import flixPostXL
from flixOptExcel.Evaluation.flixPostprocessingXL import cModelVisualizer, cVisuData
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
        co2_limit = [limit for limit in self._general_infos["CO2-limit"]]
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