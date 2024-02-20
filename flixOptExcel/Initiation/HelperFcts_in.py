# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import re
from typing import List, Literal, Union

from flixOpt.flixPlotHelperFcts import *


###############################################################################################################
# Validation
def check_dataframe_consistency(df: pd.DataFrame, years: List[int], name_of_df: str = "Unnamed Dataframe"):
    if len(df.index) / 8760 != len(years):
        raise Exception(f"Length of '{name_of_df}': {len(df)}; Number of years: {len(years)}; Doesn't match.")

    columns_with_nan = df.columns[df.isna().any()]
    if not columns_with_nan.empty:
        raise Exception(f"There are missing values in the columns: {columns_with_nan}.")


def is_valid_format_segmentsOfFlows(input_string: str, mode: Literal['validate', 'decode']) -> Union[bool, list]:
    '''
    This function was written to check if a string is of the format "0;0 ;5;10 ; 10;30"
    In mode 'validate, returns bool. In mode 'decode', returns a list of numbers
    ----------
    Returns
    -------
    bool
    '''

    # Replace commas with dots to handle decimal separators
    input_string = input_string.replace(',', '.')

    # Split the string into a list of substrings using semicolon as the delimiter
    numbers_str = input_string.split(';')
    # Convert each substring to either int or float
    numbers = [int(num) if '.' not in num else float(num) for num in numbers_str]

    if not isinstance(numbers, list):
        pass
        # raise Exception("Conversion to segmentsOfFlows didnt work. Use numbers, seperated by ';'")
    elif not all(isinstance(element, (int, float)) for element in numbers):
        pass
        # raise Exception("Conversion to segmentsOfFlows didnt work. Use numbers, seperated by ';'")
    else:
        if mode == 'validate':
            return True
        elif mode == 'decode':
            return numbers
        else:
            raise Exception(f"{mode} is not a valid mode.")
    if mode == 'validate':
        return False
    else:
        raise Exception("Error encountered in parsing of String")


def repeat_elements_of_list(original_list: [int], repetitions: int = 8760) -> np.ndarray:
    '''
    repeats each element of the list x times. If list is None, returns None
    This function was written for the creatiion of "exists"
    ----------

    Returns
    -------
    np.ndarray

    '''
    repeated_list = [item for item in original_list for _ in range(repetitions)]
    repeated_list: list
    return np.array(repeated_list)


###############################################################################################################
# calculation of Time Series
def calculate_hourly_rolling_mean(series: pd.Series, window_size: int = 24) -> pd.Series:
    """
        Calculate the hourly rolling mean of a time series.

        Parameters:
        - series (pd.Series): Time series data with hourly values. It should be indexed with datetime.
        - window_size (int): Size of the rolling window. Default is 24.

        Returns:
        - pd.Series: Hourly rolling mean of the input time series.

        Raises:
        - ValueError: If the index of the series is not in datetime format or if the hourly step is not 1 hour.

        Example:
        ```
        hourly_data = pd.Series(...)  # Replace ... with your hourly data
        result = calculate_hourly_rolling_mean(hourly_data)
        ```

        """
    # Check if the index is in datetime format
    if not pd.api.types.is_datetime64_any_dtype(series.index):
        raise ValueError("The index of the input series must be in datetime format.")

    # Check if the hourly step is 1 hour for every step
    hourly_steps = (series.index[1:] - series.index[:-1]).total_seconds() / 3600
    if not all(step == 1 for step in hourly_steps):
        raise ValueError("The time series must have a consistent 1-hour hourly step.")

    ser = series.copy()
    # Calculate the rolling mean using the specified window size
    rolling_mean = ser.rolling(window=window_size).mean()

    # Fill the missing values in 'rolling_mean' with the mean values of the series in this area
    rolling_mean.iloc[:window_size] = ser.iloc[:24].mean()

    return rolling_mean


def linear_interpolation_with_bounds(input_data: pd.Series, lower_bound: float, upper_bound: float,
                                     value_below_bound: float, value_above_bound: float) -> pd.Series:
    """
    Apply linear interpolation within specified bounds and assign fixed values outside the bounds.

    Parameters:
    - input_data (pd.Series): Input dataset.
    - lower_bound (float): Lower bound for linear interpolation.
    - upper_bound (float): Upper bound for linear interpolation.
    - value_below_bound (float): Value assigned to points below the lower bound.
    - value_above_bound (float): Value assigned to points above the upper bound.

    Returns:
    - pd.Series: New series with linear interpolation within bounds and fixed values outside.

    Example:
    ```
    # Create a sample dataset
    input_series = pd.Series([8, 12, 18, 25, 22, 30, 5, 14], index=pd.date_range('2023-01-01', periods=8, freq='D'))

    # Apply linear interpolation with bounds
    result = linear_interpolation_with_bounds(input_series, 10, 20, 5, 30)
    print(result)
    ```

    """
    output_array = np.zeros_like(input_data)
    for i in range(len(input_data)):
        if input_data.iloc[i] <= lower_bound:
            output_array[i] = value_below_bound
        elif input_data.iloc[i] >= upper_bound:
            output_array[i] = value_above_bound
        else:
            output_array[i] = (value_below_bound +
                               ((value_below_bound - value_above_bound) / (lower_bound - upper_bound)) *
                               (input_data.iloc[i] - lower_bound))
    return pd.Series(output_array, index=input_data.index)


###############################################################################################################
# Component data
def handle_component_data(df: pd.DataFrame) -> dict:
    '''
    This function was written to read the component data from an excel file
    ----------
    Returns
    -------
    '''

    # <editor-fold desc="Check for invalid Comp types">
    Erzeugertypen = ('KWK', 'Kessel', 'Speicher', 'EHK', 'Waermepumpe', 'AbwaermeHT', 'AbwaermeWP', 'Rueckkuehler',
                     'KWKekt')  # DONT CHANGE!!!
    for typ in df.iloc[0, :].dropna():
        if typ not in Erzeugertypen: raise Exception(
            f"{typ} is not an accepted type of Component. Accepted types are: {Erzeugertypen}")
    # </editor-fold>

    # <editor-fold desc="Iterate through unique values and create specific DataFrames for each type">
    # Create a dictionary to store DataFrames for each unique value
    Erzeugerdaten = {}
    for value in Erzeugertypen:
        # Select columns where the first row has the current value
        subset_df = df.loc[:, df.iloc[0] == value]

        if subset_df.shape[1] <= 1: continue  # skip, if no data inside

        # Resetting the index and droping the first column
        subset_df = subset_df.drop(0).reset_index(drop=True)

        # Rename the Columns to the Values of the first row in the created dataframe and drop the first row
        subset_df.columns = subset_df.iloc[0]
        # Rename the column at position 0
        column_names = subset_df.columns.tolist()
        column_names[0] = "category"
        subset_df.columns = column_names

        # subset_df = subset_df.drop(0).reset_index(drop=True)

        # Drop all unnecessary Rows and Cols from the dataframe
        subset_df = subset_df.dropna(axis=0, how='all').dropna(axis=1, how='all')

        # set index to the first column
        subset_df.set_index('category', inplace=True)

        # Store the subset DataFrame in the dictionary
        Erzeugerdaten[value] = subset_df
    # </editor-fold>
    print("Component Data was read successfully")

    return Erzeugerdaten


def convert_component_data_types(component_data: dict):
    '''
    This function was written to convert the component data to the right data types and do some assignments
    ----------
    :param component_data: dict of pd.Dataframes
    Returns
    -------
    dict of pd.Dataframes
    '''

    for key, subset_df in component_data.items():
        # Replace all nan values with None
        subset_df.replace({np.nan: None}, inplace=True)

        # replace "ja" and "nein" with True and False
        subset_df.replace({'ja': True, 'Ja': True, 'True': True, 'true': True,
                           'nein': False, 'Nein': False, 'false': False, 'False': False}, inplace=True)

        # check if

    return component_data


def combine_dicts_of_component_data(component_data_1, component_data_2):
    '''
    This function was written to combine the Dataframes of the different component types into one dict
    ----------
    Returns
    -------
    dict
    '''
    result_dict = {}
    for key in set(component_data_1.keys()) | set(component_data_2.keys()):
        if key in component_data_1 and key in component_data_2:
            duplicates = set(component_data_1[key].columns) & set(component_data_2[key].columns)
            if duplicates:  # if there are duplicates
                raise Exception(f"There are following Duplicates of type '{key}': {duplicates}'. Please rename them.")
            else:
                result_dict[key] = pd.concat([component_data_1[key], component_data_2[key]], axis=1)
        elif key in component_data_1:
            result_dict[key] = component_data_1[key].copy()
        elif key in component_data_2:
            result_dict[key] = component_data_2[key].copy()

    return result_dict


def convert_component_data_for_looping_through(Erzeugerdaten):
    '''
    This function was written to prepare the component data for looping through and creating the components
    It converts the component data for each type from a dataframe to a dict for each component.
    Further it removes all None values
    :param Erzeugerdaten: dict of pd.Dataframes
    ----------
    Returns
    -------
    dict of lists of dicts
    '''
    ErzDaten = {}
    for typ in Erzeugerdaten:
        ErzDaten[typ] = list()
        for comp in Erzeugerdaten[typ].columns:
            erzeugerdaten_as_dict = Erzeugerdaten[typ][comp].to_dict()
            erzeugerdaten_as_dict_wo_none = {k: v for k, v in erzeugerdaten_as_dict.items() if v is not None}
            ErzDaten[typ].append(erzeugerdaten_as_dict_wo_none)
            if not ErzDaten[typ]:  # if list is empty
                ErzDaten.pop(typ)

    return ErzDaten
