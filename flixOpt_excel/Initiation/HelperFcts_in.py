# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
from typing import List

from flixOpt.flixPlotHelperFcts import *
###############################################################################################################
# Validation
def check_dataframe_consistency(df: pd.DataFrame, years: List[int], name_of_df: str = "Unnamed Dataframe"):
    if len(df.index) / 8760 != len(years):
        raise Exception(f"Length of '{name_of_df}': {len(df)}; Number of years: {len(years)}; Doesn't match.")
    if df.isnull().any().any():
        print(f"There are missing values in '{name_of_df}'.")

def is_valid_format(input_string:str):
    '''
    This function was written to check if a string is of the format "min-max"
    ----------
    Returns
    -------
    bool
    '''
    import re
    # Define the regular expression pattern
    pattern = r'^\d+-\d+$'
    #   '^' asserts the start of the string
    #   '\d+' matches one or more digits
    #   '-' matches the hyphen character
    #   '$' asserts the end of the string

    # Use re.match to check if the string matches the pattern
    if re.match(pattern, input_string):
        return True
    else:
        return False

def to_ndarray(value, desired_length: int) -> np.ndarray:
    if isinstance(value, np.ndarray):
        ar = value
    elif isinstance(value, (pd.DataFrame, pd.Series)):
        ar = np.array(value)
    elif isinstance(value, (int, float)):
        ar = np.ones(desired_length) * value
    else:
        raise TypeError()

    if len(ar) != desired_length:
        raise Exception("length check failed")

    return ar

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

def handle_heating_network(zeitreihen: pd.DataFrame) -> pd.DataFrame:
    """
    Handle heating network parameters in the input DataFrame.

    This function calculates or checks the presence of key parameters related to the heating network,
    including supply temperature (TVL_FWN), return temperature (TRL_FWN), and network losses (SinkLossHeat).
    If not already present in the dataframe, creates them and returns the filled dataframe

    Parameters:
    - zeitreihen (pd.DataFrame): Input DataFrame containing time series data.

    Returns:
    pd.DataFrame

    Raises:
    - Exception: If one of "TVL_FWN" or "TRL_FWN" is not present in the input DataFrame and needs calculation.

    Example:
    ```python
    handle_heating_network(my_dataframe)
    ```

    """
    if "TVL_FWN" in zeitreihen.keys() and "TRL_FWN" not in zeitreihen.keys():
        raise Exception("If 'TVL_FWN' is given, 'TRL_FWN' also has to be in the Input Dataset")
    elif "TVL_FWN" not in zeitreihen.keys() and "TRL_FWN" in zeitreihen.keys():
        raise Exception("If 'TRL_FWN' is given, 'TVL_FWN' also has to be in the Input Dataset")
    elif "TVL_FWN" and "TRL_FWN" in zeitreihen.keys():
        print("TVL_FWN and TRL_FWN where included in the input data set")
    else:
        # Berechnung der Vorlauftemperatur
        zeitreihen["TVL_FWN"] = linear_interpolation_with_bounds(input_data=zeitreihen["Tamb24mean"],
                                                                 lower_bound=-9, upper_bound=11,
                                                                 value_below_bound=125, value_above_bound=105)
        # TODO: Custom Function?
        zeitreihen["TRL_FWN"] = pd.DataFrame(
            np.ones_like(zeitreihen["TVL_FWN"]) * 60, index=zeitreihen.index)

    if "sinkLossHeat" not in zeitreihen.keys():  # Berechnung der Netzverluste
        k_loss_netz = 0.4640  # in MWh/K        # Vereinfacht, ohne Berücksichtigung einer sich ändernden Netzlänge
        zeitreihen["SinkLossHeat"] = k_loss_netz * (
                (zeitreihen["TVL_FWN"] + zeitreihen["TRL_FWN"]) / 2 - zeitreihen["Tamb"])
    else:
        print("Heating losses where included in the input data set")

    return zeitreihen

def createCOPfromTS(TqTS, TsTS, eta=0.5) -> np.ndarray:
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
    #Celsius zu Kelvin
    TqTS=TqTS+273.15
    TsTS = TsTS + 273.15
    COPTS=( TsTS/(TsTS-TqTS) )    *eta
    return COPTS

def calculate_relative_capacity_of_storage(item, Zeitreihen, dT_max=65):
    '''
    This function was written to calculate the relative capacity of a Storage due to the changing
    temperature Spread in a Heating network
    ----------

    Returns
    -------
    list
    '''
    if item["calculateDT"]:
        maxReldT = ((Zeitreihen["TVL_FWN"] - Zeitreihen["TRL_FWN"]) / dT_max).values.tolist()
        maxReldT.append(maxReldT[-1])
    else:
        maxReldT = 1

    return maxReldT

def calculate_co2_credit_for_el_production(array_length, t_vl, t_rl, t_amb, n_el, n_th, co2_fuel):
    t_vl = to_ndarray(t_vl, array_length) + 273.15
    t_rl = to_ndarray(t_rl, array_length) + 273.15
    t_amb = to_ndarray(t_amb, array_length) + 273.15
    n_el = to_ndarray(n_el, array_length)
    n_th = to_ndarray(n_th, array_length)
    co2_fuel = to_ndarray(co2_fuel, array_length)
    if any(len(arg) != array_length for arg in [t_vl, t_rl, t_amb, n_el, n_th, co2_fuel]):
        raise Exception("Length check failed")
    # Berechnung der co2-Gutschrift für die Stromproduktion nach der Carnot-Methode
    t_s = (t_vl - t_rl) / np.log((t_vl / t_rl))  # Temperatur Nutzwärme
    n_carnot = 1 - (t_amb / t_s)

    a_el = (1 * n_el) / (n_el + n_carnot * n_th)
    f_el = a_el / n_el
    co2_el = f_el * co2_fuel

    # a_th = (n_carnot * n_th) / (n_el + n_carnot * n_th)
    # f_th = a_th/n_th
    # co2_th = f_th * co2_fuel

    return co2_el

###############################################################################################################
# Reading of the excel sheet
def handle_component_data(df:pd.DataFrame)-> dict:
    '''
    This function was written to read the component data from an excel file
    ----------
    Returns
    -------
    '''

    # <editor-fold desc="Check for invalid Comp types">
    Erzeugertypen = ('KWK', 'Kessel', 'Speicher', 'EHK', 'Waermepumpe', 'AbwaermeHT', 'AbwaermeWP', 'Rueckkuehler',
                     'KWKekt')  # DONT CHANGE!!!
    print("Accepted types of Components:")
    print(Erzeugertypen)
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

        if subset_df.shape[1] <= 1: continue # skip, if no data inside

        # Resetting the index and droping the first column
        subset_df = subset_df.drop(0).reset_index(drop=True)

        # Rename the Columns to the Values of the first row in the created dataframe and drop the first row
        subset_df.columns = subset_df.iloc[0]
        # Rename the column at position 0
        column_names = subset_df.columns.tolist()
        column_names[0] = "category"
        subset_df.columns = column_names

        #subset_df = subset_df.drop(0).reset_index(drop=True)

        # Drop all unnecessary Rows and Cols from the dataframe
        subset_df = subset_df.dropna(axis=0, how='all').dropna(axis=1, how='all')

        #set index to the first column
        subset_df.set_index('category', inplace=True)

        # change the labels of the dataframe
        subset_df = relabel_component_data(subset_df)

        # Store the subset DataFrame in the dictionary
        Erzeugerdaten[value] = subset_df
    # </editor-fold>
    print("Component Data was read successfully")

    return Erzeugerdaten

def relabel_component_data(df:pd.DataFrame):
    '''
    This function renames the indexes of a Dataframe by a custom mapping
    ----------
    Returns
    -------
    pd.DataFrame
    '''
    name_mapping = {'Name': 'label',
                    'Gruppe': 'group',
                    'Optional': 'optional',
                    'Thermische Leistung': 'nominal_val',
                    'exists': 'exists',
                    'eta_th': 'eta_th',
                    'eta_el': 'eta_el',
                    'Fixkosten pro Jahr': 'costs_fix',
                    'Fixkosten pro MW und Jahr': 'costs_spec',
                    'Förderung pro Jahr': 'fund_fix',
                    'Förderung pro MW und Jahr': 'fund_spec',
                    'Brennstoff': 'Brennstoff',
                    'Zusatzkosten pro MWh Brennstoff': 'ZusatzkostenEnergieInput',
                    'Zusatzkosten pro MWh Strom': 'ZusatzkostenEnergieInput',
                    'Startjahr': 'Startjahr',
                    'Endjahr': 'Endjahr',

                    # KWKekt
                    'Thermische Leistung (Stützpunkte)': 'steps_Qth',
                    'Elektrische Leistung (Stützpunkte)': 'steps_Pel',
                    'Brennstoff Leistung': 'nom_val_Qfu',

                    # Wärmepumpen
                    'MindestSCOP': 'MindestSCOP',
                    'COP': 'COP',
                    'Betriebskostenförderung BEW': 'fund_op',
                    'COP berechnen': 'calc_COP',
                    'Zeitreihe für Einsatzbeschränkung': 'TS_for_limiting_of_useage',
                    'Untergrenze für Einsatz': 'lower_limit_of_useage',
                    'Beschränkung der Leistung': 'max_rel_th',

                    # Abwaerme
                    'Abwärmekosten': 'costsPerFlowHour_abw',

                    # Speicher
                    'Kapazität [MWh]': 'capacity',
                    'Lade/Entladeleistung [MW'
                    ']': 'nominal_val',
                    'AbhängigkeitVonDT': 'calculateDT',
                    'eta_load': 'eta_load',
                    'eta_unload': 'eta_unload',
                    'VerlustProStunde': 'fracLossPerHour',
                    'Fixkosten pro MWh und Jahr': 'costs_spec_capacity',
                    'Förderung pro MWh und Jahr': 'fund_spec_capacity',

                    # Rueckkuehler
                    'Strombedarf': 'specificElectricityDemand',
                    'KostenProBetriebsstunde': 'costsPerRunningHour',
                    'Beschränkung Einsatzzeit': 'Beschränkung Einsatzzeit',

                    }

    unmapped_indexes = set(df.index) - set(name_mapping.keys())
    df.rename(index=name_mapping, inplace=True)

    if unmapped_indexes:
        raise Exception(f"In {df.columns}, There are indexes not in the mapping: {unmapped_indexes}")

    return df

###############################################################################################################
# Handling Component data
def get_invest_from_excel(item: dict, costs_effect: cEffectType, funding_effect: cEffectType, outputYears,
                          is_flow:bool=False, is_storage:bool=False, is_flow_of_storage:bool=False) -> cInvestArgs:
    '''
    This function was written for creating the cInvestArgs in an easier and more compact way
    ----------
    item : dict
        Dictionary with the data of the component. Must contain the keys:
        "nominal_val" OR "capacity", "costs_fix", "costs_spec", "fund_fix", "fund_spec", "Startjahr", "Endjahr"
    costs_effect : cEffectType
        The effect type for the costs
    funding_effect : cEffectType
        The effect type for the funding
    outputYears : list
        List of the years the calculation is done for
    is_flow : bool
        True if the component is a flow
    is_storage : bool
        True if the component is a storage
    is_flow_of_storage : bool
        True if the component is a flow of a storage

    Returns
    -------
    cInvestArgs
    if is_flow_of_storage:
        returns tuple(cInvestArgs, cEffectType)
    '''
    # default values
    min_investmentSize = 0
    max_investmentSize = 100000
    if sum([is_flow, is_storage, is_flow_of_storage]) != 1:
        raise Exception("Exactly one of the following must be True: is_flow, is_storage, is_flow_of_storage")

    # How many years is the comp in the calculation?
    Startjahr = item.get("Startjahr",None)
    Endjahr = item.get("Endjahr",None)
    if Startjahr is None or Endjahr is None:
        raise Exception("Startjahr and Endjahr must be set for " + item["label"])
    multiplier = sum([1 if Startjahr <= num <= Endjahr else 0 for num in outputYears])


    # <editor-fold desc="Fallunterscheidung">
    if is_flow: # Flows (except from storage
        if item["nominal_val"] is None:
            investmentSize_is_fixed = False
        elif isinstance(item["nominal_val"], str):
            if not is_valid_format(item["nominal_val"]):
                raise Exception("If nominal_val or capacity is passed as a string, it must be of the format 'min-max'")
            investmentSize_is_fixed = False
            min_investmentSize = float(item["nominal_val"].split("-")[0])
            max_investmentSize = float(item["nominal_val"].split("-")[1])
        elif isinstance(item["nominal_val"], (int, float)):
            investmentSize_is_fixed = True
        else:
            raise Exception(f"something went wrong creating the InvestArgs for {item['label']}")

        fixCosts = parse_dict_and_rename_and_filter(item, ["costs_fix", "fund_fix"], [costs_effect, funding_effect])
        specificCosts = parse_dict_and_rename_and_filter(item, ["costs_spec", "fund_spec"], [costs_effect, funding_effect])

    elif is_storage: # Storages
        if item["capacity"] is None:
            investmentSize_is_fixed = False
        elif isinstance(item["capacity"], str):
            investmentSize_is_fixed = False
            min_investmentSize = float(item["capacity"].split("-")[0])
            max_investmentSize = float(item["capacity"].split("-")[1])
        elif isinstance(item["capacity"], (int, float)):
            investmentSize_is_fixed = True
        else:
            raise Exception(f"something went wrong creating the InvestArgs for {item['label']}")

        fixCosts = parse_dict_and_rename_and_filter(item, ["costs_fix", "fund_fix"], [costs_effect, funding_effect])
        specificCosts = parse_dict_and_rename_and_filter(item, ["costs_spec_capacity", "fund_spec_capacity"],
                                                         [costs_effect, funding_effect])

    elif is_flow_of_storage: # Flows from storage
        if item["nominal_val"] is None:
            investmentSize_is_fixed = False
        elif isinstance(item["nominal_val"], str):
            if not is_valid_format(item["nominal_val"]):
                raise Exception("If nominal_val or capacity is passed as a string, it must be of the format 'min-max'")
            investmentSize_is_fixed = False
            min_investmentSize = float(item["nominal_val"].split("-")[0])
            max_investmentSize = float(item["nominal_val"].split("-")[1])
        elif isinstance(item["nominal_val"], (int, float)):
            investmentSize_is_fixed = True
        else:
            raise Exception(f"something went wrong creating the InvestArgs for {item['label']}")

        fixCosts = {}
        specificCosts = parse_dict_and_rename_and_filter(item, ["costs_spec", "fund_spec"], [costs_effect, funding_effect])

        multiplier = multiplier * 0.5 # because investment is split between the input and output flow
    else:
        raise Exception("Exactly one of the following must be True: is_flow, is_storage, is_flow_of_storage")
    # </editor-fold>

    # Optionales Investment ?
    if item.get("optional", None) is None:
        investment_is_optional =False
    else:
        investment_is_optional = item["optional"]

    # Multiply the costs with the number of years the comp is in the calculation
    for key in fixCosts:
        fixCosts[key] *= multiplier
    for key in specificCosts:
        specificCosts[key] *= multiplier

    Invest = cInvestArgs(fixCosts=fixCosts, specificCosts=specificCosts,
                         investmentSize_is_fixed = investmentSize_is_fixed,
                         investment_is_optional = investment_is_optional,
                         min_investmentSize=min_investmentSize, max_investmentSize=max_investmentSize)
    return Invest

def handle_operation_years(item:dict, outputYears:list) -> np.ndarray:
    '''
    This function was written to match the operation years of a component to the operation years of the system
    ----------

    Returns
    -------
    np.ndarray
    '''
    Startjahr = item.get("Startjahr",None)
    Endjahr = item.get("Endjahr",None)

    if Startjahr is None or Endjahr is None:
        raise Exception("Startjahr and Endjahr must be set for " + item["label"])

    # Create a new list with 1s and 0s based on the conditions
    list_to_repeat = [1 if Startjahr <= num <= Endjahr else 0 for num in outputYears]

    if len(list_to_repeat) == sum(list_to_repeat):
        return 1
    else:
        return np.array(repeat_elements_of_list(list_to_repeat))

def handle_nom_val(value_nom_val_or_cap):
    '''
    This functions handles the value of nom_val. This is necessary for the creation of bounds for the investment, which are passed as a string.

    '''
    if isinstance(value_nom_val_or_cap, str):
        if is_valid_format(value_nom_val_or_cap):
            return None
        else:
            raise Exception("If nominal_val or capacity is passed as a string, it must be of the format 'min-max'")
    elif value_nom_val_or_cap is None:
        return None
    elif isinstance(value_nom_val_or_cap, (int, float)):
        return value_nom_val_or_cap
    else:
        raise Exception("Wrong datatype for nominal_val or capacity. Must be int, float string of form 'min-max' or None")

def parse_dict_and_rename_and_filter(original_dict: dict, old_keys: list, new_keys: list) -> dict:
    '''
    This function was written for creating the dict for the invest costs
    ----------

    Returns
    -------
    dict

    '''

    filtered_dict = {key: original_dict.get(key, None) for key in old_keys}
    new_dict = {new_key: filtered_dict[old_key] for new_key, old_key in zip(new_keys, filtered_dict) if
                filtered_dict[old_key] not in (None, 0)}
    return new_dict

def handle_fuel_input_switch_excel(item, Preiszeitreihen,
                                   bus_erdgas, bus_h2, bus_ebs,
                                   costs, CO2, tCO2perMWhGas) -> (cBus, dict):
    '''
    This function was written to assign the right bus and fuel cost
    ----------

    Returns
    -------
    (fuel_bus: cBus, fuel_costs: dict)
    '''
    if item.get("ZusatzkostenEnergieInput") is None:
        extra_costs=0
    else:
        extra_costs = item["ZusatzkostenEnergieInput"]
    if item["Brennstoff"] == "Erdgas":
        fuel_bus = bus_erdgas
        fuel_costs = {costs: Preiszeitreihen["costsBezugGas"] + extra_costs, CO2: tCO2perMWhGas}
    elif item["Brennstoff"] == "Wasserstoff":
        fuel_bus = bus_h2
        fuel_costs = {costs: Preiszeitreihen["costsBezugH2"] + extra_costs}
    elif item["Brennstoff"] == "EBS":
        fuel_bus = bus_ebs
        fuel_costs = {costs: Preiszeitreihen["costsBezugEBS"] + extra_costs}
    else:
        raise Exception("Brennstoff '" + item["Brennstoff"] + "' is not yet implemented in Function")


    return fuel_bus, fuel_costs


def handle_COP_calculation(item, Zeitreihen, eta=0.5)-> cTSraw:
    '''
    This function was written to assign a COP to a Heat Pump
    ----------

    Returns
    -------
    (fuel_bus: cBus, fuel_costs: dict)
    '''
    # Wenn fixer COP übergeben wird
    if isinstance(item["COP"], (int, float)):
        COP = item["COP"]
    elif item["COP"] in Zeitreihen.keys():  # Wenn verlinkung zu Temperatur der waermequelle vorgegeben ist
        if item["calc_COP"]:
            COP = createCOPfromTS(TqTS=Zeitreihen[item["COP"]], TsTS=Zeitreihen["TVL_FWN"], eta=eta)
        else:
            COP = Zeitreihen[item["COP"]]
        Zeitreihen["COP" + item["label"]] = COP
        COP = cTSraw(COP)
    else:
        raise Exception("Verlinkung zwischen COP der WP " + item[
            "label"] + " und der Zeitreihe ist fehlgeschlagen. Prüfe den Namen der Zeitreihe")

    return COP

def limit_useage(item, zeitreihen)-> np.ndarray:
    '''
    Limit useage of a heat pump by a temperature bound
    Args:
        item:
        zeitreihen:

    Returns:
        np.ndarray:
    '''

    if item.get("lower_limit_of_useage") is None:
        max_rel = 1
    elif isinstance(item["lower_limit_of_useage"], (int,float)):
        if item.get("TS_for_limiting_of_useage") is None:
           raise Exception("If you want to limit the useage of a Heat Pump, you have to specify the TS to calculate the useage")
        elif item["TS_for_limiting_of_useage"] not in zeitreihen.keys():
            raise Exception(f"The specified TS {item['TS_for_limiting_of_useage']} is not in 'Zeitreihen'")
        else:
            ts_for_limiting = zeitreihen[item["TS_for_limiting_of_useage"]]
            # Create a new array based on the condition
            max_rel = (ts_for_limiting >= item["lower_limit_of_useage"]).astype(int)
    else:
        raise Exception("if you want to limit the useage of a Heat Pump, choose a number as the lower limit")

    return max_rel.values


def handle_operation_fund_of_heatpump(item, funding, COP, costs_for_electricity, fact=92)-> dict:
    '''
    This function was written to calculate the operation funding of a Heat Pump (BEW)

    ----------

    Returns
    -------
    dict
    '''
    # Betriebskostenförderung, Beschränkt auf 90% der Stromkosten
    if item.get("fund_op"):
        # Förderung pro MW_th
        ar = (COP.value - 1 / COP.value) * fact

        # Stromkosten pro MW_th
        cost_for_el_per_MWth = costs_for_electricity / COP.value

        # Begrenzung der Förderung auf 90% der Stromkosten
        ar = np.where(ar > cost_for_el_per_MWth * 0.9, cost_for_el_per_MWth * 0.9,ar)
        fund_op = {funding: ar}
    else:
        fund_op = None

    return fund_op

def convert_component_data_types(component_data:dict):
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
        subset_df.replace({'ja': True, 'nein': False}, inplace=True)

        # check if

    return component_data

def combine_dicts_of_component_data(component_data_1,component_data_2):
    '''
    This function was written to combine the Dataframes of the different component types into one dict
    ----------
    Returns
    -------
    dict
    '''
    result_dict={}
    for key in set(component_data_1.keys()) | set(component_data_2.keys()):
        if key in component_data_1 and key in component_data_2:
            duplicates = set(component_data_1[key].columns) & set(component_data_2[key].columns)
            if duplicates: # if there are duplicates
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
    :param Erzeugerdaten: dict of pd.Dataframes
    ----------
    Returns
    -------
    dict
    '''
    ErzDaten = {}
    for typ in Erzeugerdaten:
        ErzDaten[typ] = list()
        for comp in Erzeugerdaten[typ].columns:
            ErzDaten[typ].append(Erzeugerdaten[typ][comp].to_dict())
            if not ErzDaten[typ]: # if list is empty
                ErzDaten.pop(typ)

    return ErzDaten

def repeat_elements_of_list(original_list:[int], repetitions:int=8760) -> np.ndarray:
    '''
    repeats each element of the list x times. If list is None, returns None
    This function was written for the creatiion of "exists"
    ----------

    Returns
    -------
    np.ndarray

    '''
    if original_list is None: return None
    repeated_array = [item for item in original_list for _ in range(repetitions)]
    return repeated_array


def string_to_list(delimited_string: str, delimiter: str = '-') -> list:
    """
    Convert a string of hyphen-separated numbers to a list of floats.

    Parameters:
    - delimited_string (str): The input string containing delimited numbers.
    - delimiter (str): The delimiter

    Returns:
    - list: A list of floats representing the numbers in the input string.
    """
    return list(map(float, delimited_string.split(delimiter)))