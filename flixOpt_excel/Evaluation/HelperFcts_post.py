import pandas as pd
import numpy as np
import timeit
import datetime
import os.path
from typing import Literal, List, Union

from flixOpt_excel.Evaluation.flixPostprocessingXL import flixPostXL


def resample_data(data_frame: Union[pd.DataFrame, np.ndarray], target_years: List[int], resampling_by: Literal["Y", "D", "H"],
                   resampling_method: Literal["sum", "mean", "min","max"], initial_sampling_rate: str = "H") -> pd.DataFrame:
    '''
    Parameters
    ----------
    data_frame : Union[pd.DataFrame, np.ndarray]
        DataFrame or array containing data. Number of rows must match the initial sampling rate (safety check):
        8760 ("H") (default) or 365 ("D") per year.
    target_years : List[int]
        Target years for the new index of the DataFrame
    resampling_by : str
        "H" for hourly resampling
        "D" for daily resampling
        "Y" for yearly resampling
    resampling_method : str
        "mean" for mean value
        "sum" for sum value
        "max" for max value
        "min" for min value
    initial_sampling_rate : str
        "H" for hourly data (8760 values per year)
        "D" for daily data (365 values per year)

    Returns
    -------
    pd.DataFrame
    '''
    df = pd.DataFrame(data_frame)
    df.index = range(len(df))  # reset index

    if len(df)/8760 == len(target_years) and initial_sampling_rate == "H":
        length_per_year = 8760
    elif len(df)/365 == len(target_years) and initial_sampling_rate == "D":
        length_per_year = 365
    else:
        raise ValueError("length of dataframe and initial_sampling_rate must match: "
                         "8760 rows/year ('H') or 365 rows/year 'D'.")

    if not isinstance(target_years, list):
        target_years = [target_years]

    # create new TS for resampling, without the 29. February (filtering leap years)
    for i, year in enumerate(target_years):
        dt = pd.date_range(start=f'1/1/{year}', end=f'01/01/{year + 1}', freq=initial_sampling_rate)[:-1]
        dt = dt[~((dt.month == 2) & (dt.day == 29))]  # Remove leap year days
        df.loc[i * length_per_year:(i + 1) * length_per_year - 1, 'Timestamp'] = dt
    df = df.set_index('Timestamp')

    if resampling_method == "sum":
        df = df.resample(resampling_by).sum()
    elif resampling_method == "mean":
        df = df.resample(resampling_by).mean()
    elif resampling_method == "min":
        df = df.resample(resampling_by).min()
    elif resampling_method == "max":
        df = df.resample(resampling_by).max()
    else:
        raise ValueError("Invalid resampling method")

    # Drop all rows that aren't in the years specified in target_years
    lst = [row for row in df.index if row.year not in target_years]
    df = df.drop(index=lst)

    df = df.loc[~((df.index.month == 2) & (df.index.day == 29))]  # Remove leap year days again

    if resampling_by == "Y":
        df = df.set_index(df.index.year)  # setting the index to the plain year. No datetime anymore

    return df

def rs_in_two_steps(data_frame: Union[pd.DataFrame, np.ndarray], target_years: List[int], resampling_by: Literal["D", "Y"],
                    initial_sampling_rate: str = "H") -> pd.DataFrame:
    '''
    Parameters
    ----------
    data_frame : Union[pd.DataFrame, np.ndarray]
        DataFrame or array containing data. Number of rows must match the initial sampling rate (safety check):
        8760 ("H") (default) or 365 ("D") per year.
    target_years : List[int]
        Years for resampling
    resampling_by : str
        "D" for daily resampling
        "Y" for yearly resampling
    initial_sampling_rate : str
        "H" for hourly data
        "D" for daily data
    Returns
    -------
    pd.DataFrame
        Resampled DataFrame with new columns:
        ["Tagesmittel", "Minimum (Stunde)", "Maximum (Stunde)"]
        or new Columns:
        ["Jahresmittel", "Minimum (Tagesmittel)", "Maximum (Tagesmittel)"],
        depending on chosen "resampling_by"
    '''

    # Determine base resampling method and new columns based on resampling_by
    if resampling_by == "D":
        rs_method_base = "H"
        new_columns = ["Tagesmittel", "Minimum (Stunde)", "Maximum (Stunde)"]
    elif resampling_by == "Y":
        rs_method_base = "D"
        new_columns = ["Jahresmittel", "Minimum (Tagesmittel)", "Maximum (Tagesmittel)"]
    else:
        raise ValueError("Invalid value for resampling_by. Use 'D' for daily or 'Y' for yearly.")


    # Base resampling
    df_resampled_base = resample_data(data_frame, target_years, rs_method_base, "mean", initial_sampling_rate)

    # Resample for min, max, and mean
    min_y = resample_data(df_resampled_base, target_years, resampling_by, "min", rs_method_base)
    max_y = resample_data(df_resampled_base, target_years, resampling_by, "max", rs_method_base)
    mean_y = resample_data(df_resampled_base, target_years, resampling_by, "mean", rs_method_base)

    # Concatenate results
    df_result = pd.concat([mean_y, min_y, max_y], axis=1)
    df_result.columns = new_columns

    return df_result

def getFuelCosts(calc:flixPostXL) -> pd.DataFrame:
    '''
    Returns the costs per flow hour of every medium in a DataFrame. Data saved in a special component ("HelperPreise").

    Parameters
    ----------
    calc : flixPostXL
        Solved calculation of type flixPostXL.

    Returns
    -------
    pd.DataFrame
        DataFrame containing the costs per flow hour for each medium. Columns represent different media,
        and rows represent the time series.
    '''
    (discard, flows) = calc.getFlowsOf("HelperPreise")
    result_dataframe = pd.DataFrame(index=calc.timeSeries)
    for flow in flows:
        name = flow.label_full.split("_")[-1]
        ar = flow.results["costsPerFlowHour_standard"]
        if isinstance(ar,(float,int)):
            ar=ar * np.ones(len(calc.timeSeries))

        new_dataframe = pd.DataFrame({name: ar}, index=calc.timeSeries)
        result_dataframe = pd.concat([result_dataframe, new_dataframe], axis=1)

    return result_dataframe.head(len(calc.timeSeries))

def reorder_columns(df:pd.DataFrame, not_sorted_columns: List[str] = None):
    '''
    Order a DataFrame by a custom function, excluding specified columns from sorting, and setting them as the first columns.

    Parameters
    ----------
    df : pd.DataFrame
        Input DataFrame.
    not_sorted_columns : List[str], optional
        Columns to exclude from sorting and set as the first columns, by default None.

    Returns
    -------
    pd.DataFrame
        DataFrame with the desired column order.
    '''
    if isinstance(df, pd.Series): df = df.to_frame().T

    means = df.sum()
    sorted_columns = means.sort_values(ascending=False).index
    sorted_df = df[sorted_columns]

    # Select the remaining columns excluding the first two
    if not_sorted_columns is None:
        other_columns = [col for col in sorted_df.columns]
        # Create a new DataFrame with the desired column order
        new_order_df = sorted_df[other_columns]
    else:
        other_columns = [col for col in sorted_df.columns if col not in not_sorted_columns]

        # Create a new DataFrame with the desired column order
        new_order_df = pd.concat([sorted_df[not_sorted_columns], sorted_df[other_columns]], axis=1)

    return new_order_df


#old
def sum_columns_by_prefix(df:pd.DataFrame, prefixes:List[str]):
    """
        Sum up columns in a DataFrame based on specified prefixes.

        Parameters
        ----------
        df : pd.DataFrame
            Pandas DataFrame.
        prefixes : List[str]
            List of strings, prefixes to filter columns.

        Returns
        -------
        pd.DataFrame
            DataFrame with the summation results for each prefix, and the other, not summed columns.
        """
    result_columns = []

    # Track columns that were already considered for summation
    columns_already_summed = set()

    for prefix in prefixes:
        selected_columns = df.filter(regex=f'^{prefix}', axis=1) # DataFrame
        selected_columns_sum = selected_columns.sum(axis=1).rename(prefix) # Dataframe
        if selected_columns.empty:
            continue
        else:
            result_columns.append(selected_columns_sum)
            # Update the set of columns already summed
            columns_already_summed.update(selected_columns.columns)

    # Include only the columns not previously summed up
    non_summed_columns = df.columns.difference(columns_already_summed)
    if len(result_columns) == 0: # if list is  empty
        result_df = df[non_summed_columns]
    else:
        result_df = pd.concat(result_columns, axis=1)
        result_df = pd.concat([result_df, df[non_summed_columns]], axis=1)


    return result_df




class RuntimeTracker():
    def __init__(self, label:str):
        '''
        Parameters
        ----------
        label : str
            name of Tracker
        group : int
            0: general
            1: Year 1
            2: Year 2
            ...

        Returns
        -------
        None.

        '''
        self.label=label
        self.startTime = timeit.default_timer()
        self.stop_time = None

    def stop(self):
        if self.stop_time is None:
            self.stop_time = timeit.default_timer()
            self.text()
        else:
            print("Timer was already stopped")

    def duration(self):
        self.stop()
        return self.stop_time - self.startTime


    def text(self, supress_print=False):
        if self.stop_time is None: raise Exception(f"Timer '{self.label} not stopped yet")
        duration = datetime.timedelta(milliseconds=(self.stop_time - self.startTime)*1000)
        text=f"{duration}  [HH:MM:SS.ms]: Runtime of {self.label}"
        if not supress_print:
            print(text)
        return text


print()