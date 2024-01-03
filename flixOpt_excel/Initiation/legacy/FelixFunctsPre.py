# -*- coding: utf-8 -*-
"""
Created on Thu Jun 16 11:19:17 2022
developed by Felix Panitz* and Peter Stange*
* at Chair of Building Energy Systems and Heat Supply, Technische Universität Dresden
"""
import copy

import pandas as pd
import numpy as np

from flixComps import *
from OwnModels.flixOpt_FB.FelixFuncts import *
from flixPlotHelperFcts import *
import ast
import os

# all Functions for Preprocessing Data
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

def createSTprofileDresdenTRY(Tvl:np.ndarray =120, Trl:np.ndarray=90, KollType:str="VRK") -> pd.Series:
    '''
    Calculates the resources from ST for the weather Data from TRY 2015 in Dresden Johannstadt
    Takes Tvl and Trl of the heating network into account

    Parameters
    ----------
    Tvl : np.ndarray
        Vorlauftemperatur des FWN
    Trl : np.ndarray
        Rücklauftemperatur des FWN

    Returns
    -------
    Q_MW_per_MWp: pd.Series
        Erzeugung als Zeitreihe in MWh/h  pro MWp (Berechnet aus Effizienz der Anlage)

    -------
    Comments:
    # Inputs: Tvl, Trl          [optimal: , Tamb + Strahlungsdaten; macht nur zusammen sinn, sonst verfälschung der wetterdaten.
    # Wenn nur Tvl/Trl übergeben werden, dann wird ein realistisches Wetterprofil für die Berechnung genutzt

    # resources: Erzeugung als Einspeisung/kWp über gesamten Optimierungszeitraum

    '''
    print("################### Calculating ST-Profile ##################")
    import PySolCalc.components as comp
    import PySolCalc.irradiance as irr
    from PySolCalc.pipe import Pipes

    # <editor-fold desc="Standort und Wetterdaten">
    ############ Standort ##########
    # Koordinaten TRY (Dresden Johannstadt)
    long = 13.7750
    lat = 51.0504
    base = os.getcwd()
    if "\\" in base:
        path="/".join(os.getcwd().split("\\")[:-1])
    elif "/" in base:
        path = "/".join(os.getcwd().split("/")[:-1])
    pathTRY = path +"/PySolCalc/FB_test/" \
              "product_wgs84_20220506__115234/TRY_510504137750/TRY2015_510504137750_Jahr.dat"
    # </editor-fold>

    # <editor-fold desc="Modeling Options">
    source = "dwd_try"
    sky_model = "T/C/K"  # "hay_and_davies"
    # </editor-fold>

    # <editor-fold desc="Solarthermie-Feld">
    # Collectorparameter
    col_height = 2.03
    # TODO: Typische Kollektoren für FW finden (VRK und FK)
    if KollType=="VRK":
        col_model = "XL 19/49" # passt, ist Typisch
        peakPowerPerSqm50K= 580 # W/m^2 bei Standarttestbedingungen (EU)
    elif KollType=="FK":
        col_model = "GK3102 M/PR"
        peakPowerPerSqm50K = 580 # W/m^2 bei Standarttestbedingungen (EU)
    else:
        raise Exception(KollType +" is not a valid collector type")

    # Solarfeldparameter
    ST_tilt = 20
    ST_azimuth = 0
    ST_l_between_rows = 3.5 # Wert von ST-Anlage TUD
    ST_module_area = (1 / (peakPowerPerSqm50K/1000))   *1000*5  #    = m^2/ (5 MW)

    # </editor-fold>

    ##################### Ab hier nur noch berechnung ########################
    weather = irr.cWeather(latitude=lat,
                           longitude=long,
                           path=pathTRY,
                           source=source,
                           # timeframe_start=zeit_start,
                           # timeframe_end=zeit_ende,
                           # interpolate=4
                           )

    # <editor-fold desc="Modelling of Solarthermie">

    colST = comp.cSTCollector("ST1", height=col_height, name=col_model)
    # calculate necessary Module area for a peak Power of 5 MW
    #                (1 / (     W/m^2             /1000))   *1000 = (    m^2/kW   )* 1000 =      m^2/MW

    solarFieldST = comp.cSolarField(colST, tilt=ST_tilt, azimuth=ST_azimuth,
                                    module_area=ST_module_area, l_between_rows=ST_l_between_rows)

    irradianceST = irr.Irradiance(weather=weather, collector_slope=solarFieldST.tilt,
                                  collector_azimuth=solarFieldST.azimuth, sky_model=sky_model)
    # Standardwerte aus Solites Excel
    pipes = Pipes(lambda_d=0.035, lambda_e=2, laying_depth=0.8,
                  internal_pipe_volume=3, internal_pipe_loss_factor=0.06,
                  binding_pipe_length=20, binding_pipe_diameter=0.207,
                  binding_pipe_loss_factor=2.5, constant_ambient_temperature=10,
                  piping_type=3, c_fluid=4180, rho_fluid=1000)

    colFieldST = comp.cCollField(solarField=solarFieldST, irr=irradianceST)
    #PPeakOfField = colFieldST.solarField.peak_power / 1000 / colFieldST.solarField.module_area

    one_step_result = colFieldST.calculate_Solites_step(step=0, T_ff=Tvl[0], T_rf=Trl[0], pipes=pipes, timestep=3600)

    TS_result_ST = colFieldST.calculate_Solites_model(T_ff=Tvl, T_rf=Trl, pipes=pipes, timestep=3600)
    Q_MW= TS_result_ST.Q_kW/1000 # Erzeugung pro Stunde pro 1000m^2
    Q_MW_per_MWp=Q_MW/5
    #Q_MW_per_1000m2 = TS_result_ST.Q_kW/1000 # Erzeugung pro Stunde pro 1000m^2

    return Q_MW_per_MWp

# For Tvl FWN
def linearEqWithBounds( inLow:float= -9, inUp:float= 11,
                        outLow:float = 125, outUp:float = 105, in_Series:pd.Series=None):
    '''
    Creates a new array based on the input array. In between low_bound and upper_bound a linear relation is applied

    Parameters
    ----------
    inLowerBound : float
        For values below this value, the resources contains the value of outLowerBound
    inUpperBound : float
        For values above this value, the resources contains the value of outUpperBound
    outLowerBound : float

    outUpperBound : float

    Returns
    -------
    output: pd.Series
        new pd.Series

    -------

    '''

    # Berechnung der Vorlauftemperatur TODO: Wie externe Eingabe (Excel)?
    # DIe Vorlauftemperatur ist zwischen Tmin und Tmax abhängig von der außentemperatur (linear)
    # Außerhalb ist sie gleich einer fixen tmeperatur (TvlMax, TvlMin)

    output_array = np.zeros_like(in_Series)
    for i in range(len(in_Series)):
        if in_Series[i]<=inLow:
            output_array[i]=outLow
        elif in_Series[i]>=inUp:
            output_array[i]=outUp
        else: # Hier muss der lineare zusammenhang rein
            output_array[i] =\
                outLow + (   (outLow-outUp) / (inLow-inUp)   ) * ( in_Series[i] - inLow )
                #  120 + (   (120-90)       / (-12-5)        ) * ( -8           - -12 )
                #  120 + (   (30)           / (-17)          ) * ( 4                  )
                #  120 + (   -1,8                            ) * ( 4                  )
                #  120 + (   -1,8                            ) * ( 4                  )
                #  120 + (   -1,8                              * ( 4                  )
    return pd.Series(output_array)

############################################################################################################
# <editor-fold desc="Processing of the ErzDat Excel">
def cleanUpErzDict(d:dict)->dict:
    '''
    This functions deletes unnecessary keys from the dict
    ----------
    d : dict
        Input dict with Data of Components

    Returns
    -------
    dict

    '''
    # delete the unnecessary key/s (columns = first level) = Comps
    delete = [Comp for Comp in d if (Comp.startswith('-') or Comp.startswith('Unnamed') or Comp.startswith('Type'))]
    for Comp in delete:
        del d[Comp]

    # delete the unnecessary key/s (Types - second level) = Types
    for Comp in d:
        delete = [Type for Type in d[Comp] if (Type.startswith('-') or Type.startswith('Unnamed'))]
        for Type in delete:
            del d[Comp][Type]

    # delete the unnecessary key/s (VarName - third level)
    for Comp in d:
        for Type in d[Comp]:
            delete = [VarName for VarName in d[Comp][Type] if (d[Comp][Type][VarName]=='-')]
            for VarName in delete:
                del d[Comp][Type][VarName]

    # Delete unnecessary Types (FlowIn1... FlowIn2..., NewParam) if no keys are existing
    for Comp in d:
        delete = [Type for Type in d[Comp] if (len(d[Comp][Type].keys())==0 )]
        for Type in delete:
            del d[Comp][Type]
            
    return d

def convertDataTypes(d:dict)->dict:
    '''
    This functions converts string of "True"/"False" and "None" to corresponding data Types
    ----------
    d : dict
        Input dict with Data of Components

    Returns
    -------
    dict
    '''
    #Convert string to bool for every Comp and every Type and every VarName
    for Comp in d:
        for Type in d[Comp]:
            for VarName in d[Comp][Type]:
                if d[Comp][Type][VarName]=="True":
                    d[Comp][Type][VarName]=True
                elif d[Comp][Type][VarName]=="False":
                    d[Comp][Type][VarName]=False

    #Convert string to None for every Comp and every Type and every VarName
    for Comp in d:
        for Type in d[Comp]:
            for VarName in d[Comp][Type]:
                if d[Comp][Type][VarName]=="None":
                    d[Comp][Type][VarName]=None

    # TODO: Convert the segmentsOfFlows to a useable From
    for Comp in d:
        if "Segments" in d[Comp]:
            for VarName in d[Comp]["Segments"]:
                lstStr=d[Comp]["Segments"][VarName]
                d[Comp]["Segments"][VarName]=ast.literal_eval(lstStr)
                print("########")
                print(lstStr)
                print(d[Comp]["Segments"][VarName])
                print(type(d[Comp]["Segments"][VarName]))
    return d

def applyCostsperFlowHour(d:dict, InDF,costs,CO2,CO2faktorGas,CO2FaktorAbfall)->dict:
    '''
    This functions applies the corresponding costs of fuels/electr to the Flows
    Also Rewards
    ----------
    d : dict
        Input dict with Data of Components
    InDict : dict
        Input Dict containing the Cost of the Flow

    Returns
    -------
    dict
    '''
    for Comp in d:
        for Type in d[Comp]:
            if Type.startswith("FlowIn"):
                if d[Comp][Type]["bus"] == "Fernwaerme":
                    raise Exception("No Costs Implemented for FW Input")
                elif d[Comp][Type]["bus"] == "StromBez":
                    d[Comp][Type]["costsPerFlowHour"] = InDF["costsBezugEl"]
                elif d[Comp][Type]["bus"] == "Gas":
                    d[Comp][Type]["costsPerFlowHour"] = {costs:InDF["costsBezugGas"], CO2:CO2faktorGas}
                elif d[Comp][Type]["bus"] == "H2":
                    d[Comp][Type]["costsPerFlowHour"] = InDF["costsBezugH2"]
                    raise Exception("No costs for InFlow H2 implemented yet")
                elif d[Comp][Type]["bus"] == "Abfall":
                    d[Comp][Type]["costsPerFlowHour"] = {costs: InDF["costsBezugAbfall"], CO2: CO2FaktorAbfall}
                elif d[Comp][Type]["bus"] == "StromEinsp":
                    raise Exception("Dont use StromEinsp as Input")
            if Type.startswith("FlowOut"):
                if d[Comp][Type]["bus"] == "Fernwaerme":
                    raise Exception("No Costs Implemented for FW resources")
                elif d[Comp][Type]["bus"] == "StromEinsp":
                    d[Comp][Type]["costsPerFlowHour"] = InDF["costsEinspEl"]
                elif d[Comp][Type]["bus"] == "Gas":
                    raise Exception("No costs for OutFlow of Gas implemented yet")
                elif d[Comp][Type]["bus"] == "H2":
                    raise Exception("No costs for OutFlow H2 implemented yet")
                elif d[Comp][Type]["bus"] == "Abfall":
                    raise Exception("No costs for OutFlow of Abfall implemented yet")
                elif d[Comp][Type]["bus"] == "StromBez":
                    raise Exception("Dont use StromBez as Input")
    return d

def applyMaxRel(d:dict, listOfYears:list, lenPerYear:int=8760)->dict:
    '''
    This functions creates a TS out of min_rel and max_rel, that corresponds to the existance of the component.
    It applies it to all flows that are attached to the component
    ----------
    d : dict
        Input dict with Data of Components

    Returns
    -------
    dict
    '''
    for Comp in d:
        for Type in d[Comp]:
            if Type.startswith("Flow"):
                #Basiswerte für min_rel max_rel
                min_rel_inp = 0
                max_rel_inp = 1
                start=getStartIndex(listOfYears, d[Comp]["NewParam"]["operationStart"])
                end=getEndIndex(listOfYears, d[Comp]["NewParam"]["operationEnd"])

                #min_rel
                if 'min_rel' in d[Comp][Type]:
                    min_rel_inp = d[Comp][Type]["min_rel"] #übernehmen des Inputs, falls vorhanden
                min_rel = np.concatenate(
                    [np.zeros(start * lenPerYear), np.ones((end + 1 - start) * lenPerYear) * min_rel_inp]).tolist()
                d[Comp][Type]['min_rel']=min_rel

                #max_rel
                if 'max_rel' in d[Comp][Type]:
                    max_rel_inp = d[Comp][Type]["max_rel"] #übernehmen des Inputs, falls vorhanden
                max_rel = np.concatenate(
                    [np.zeros(start * lenPerYear), np.ones((end + 1 - start) * lenPerYear) * max_rel_inp]).tolist()
                d[Comp][Type]['max_rel']=max_rel

    return d

def applyVal_rel(d:dict, listOfYears:list, TSdict:dict, lenPerYear:int=8760,)->dict:
    '''
    This functions applies the TS for ST and =0 where ST doesn't exist yet/anymore
    ----------
    d : dict
        Input dict with Data of Components
    TotalDict : dict
        For TS of ST

    Returns
    -------
    dict
    '''
    for Comp in d:
        if Comp.startswith("ST"):
            # Grab ST TimeSeries
            if d[Comp]["ST"]["kollType"]=="VRK":
                val_rel = TSdict["STProfileVRK"]
            elif d[Comp]["ST"]["kollType"]=="FK":
                val_rel = TSdict["STProfileFK"]
            else:
                raise Exception(
                    "Als Wert für Kollektorytp in " + d[Comp] + " sind nur {VRK, FK} zulässig")

            # Gleich null setzten solange ST-Anlage nicht existiert
            start = getStartIndex(listOfYears, d[Comp]["NewParam"]["operationStart"])
            end = getEndIndex(listOfYears, d[Comp]["NewParam"]["operationEnd"])

            val_rel.iloc[0:start * lenPerYear] = 0
            if end + 1 < len(listOfYears):
                val_rel.iloc[(end + 1) * lenPerYear:len(listOfYears) * lenPerYear] = 0

            d[Comp]["FlowOut1"]["val_rel"] = val_rel

    return d

def applyFracLossStorage(d:dict, listOfYears:list, lenPerYear:int=8760,)->dict:
    '''
    This functions changes the FracLoss to a TS, that is 0 before and after the Storage doesnt exist
    ----------
    d : dict
        Input dict with Data of Components
    TotalDict : dict
        For TS of ST

    Returns
    -------
    dict
    '''
    for Comp in d:
        if Comp.startswith("Speicher"):
            # Grab ST TimeSeries
            if "fracLossPerHour" in d[Comp]["Comp"]: # If not: standard-Value = 0, so no need for creation of a TS
                fracOrg=d[Comp]["Comp"]["fracLossPerHour"]
                fracTS=np.ones(len(listOfYears)*lenPerYear)*fracOrg # create a TS

                # Gleich null setzten solange der Speicher nicht existiert
                start = getStartIndex(listOfYears, d[Comp]["NewParam"]["operationStart"])
                end = getEndIndex(listOfYears, d[Comp]["NewParam"]["operationEnd"])

                fracTS[0:start * lenPerYear] = 0
                if end + 1 < len(listOfYears):
                    fracTS[(end + 1) * lenPerYear:len(listOfYears) * lenPerYear] = 0

                d[Comp]["Comp"]["fracLossPerHour"] = fracTS

    return d

def checkPlausibilityOneInput(d:dict):
    # Test if Investment size is fixed=False; If so, nominal Value has to be None, raise exception
    # TODO Exception for Speicher, da nicht nominal_val von FlowOut optimiert wird
    # TODO: Auf neues Inputfile anpassen
    for Comp in d:
        if not Comp.startswith("Speicher"):
            if "Invest" in d[Comp]:
                if "investmentSize_is_fixed" in d[Comp]["Invest"]:
                    if d[Comp]["Invest"]["investmentSize_is_fixed"] == False:
                        if "nominal_val" in d[Comp]["FlowOut1"].keys(): #Unnötige Abfrage?
                            if d[Comp]["FlowOut1"]["nominal_val"] != None:
                                raise Exception("Nominal Value of the FlowOut1 of " + str(Comp) +
                                                " has to be changed to None, because Investment size is 'None'")
            else:
                raise Exception(
                    str(Comp) + " enthält keine Investkosten. Bitte eins ausfüllen.")
# </editor-fold>

############################################################################################################
# <editor-fold desc="InvestCosts">

#Annuise Costs
def get_a(year, q):
    return (q ** year * (q - 1)) / (q ** year - 1)

def applyInvestFunding(d):


    for Comp in d:
        # Get the Amount of funding of the Component
        fund=d["NewParam"]["investFund"]

        for Type in d[Comp]:
            if Type.startswith('Invest'):
                for VarName in d[Comp][Type]:
                    if VarName in ["fixCosts", "divestCosts", "specificCosts"]: # TODO:Funding of other costs?
                        d[Comp][Type][VarName] = d[Comp][Type][VarName] * (1-fund)
    return d

def applyYearsOfExistance(d:dict, listOfYears:list):
    '''
    This functions calculates the annual Costs from the Input Dict.
    The annual costs get multiplied with the amount of years the Component exists in the calculation,
    resulting in the Total Costs over al chosen years
    ----------
    d : dict
        Input dict with Data of Components
    InDict : dict
        Input Dict containing the Cost of the Flow

    Returns
    -------
    d : dict
    '''
    # TODO: Für costsInInvestSizeSegments implementieren
    for Comp in d:
        # Get the Operation Time of the Component
        start=d[Comp]["NewParam"]["operationStart"]
        end=d[Comp]["NewParam"]["operationEnd"]
        y = end - start # Anzahl der Jahre, die die Anlage Existiert

        # Anzahl der Jahre, die Die Anlage im Modell existiert (anzahl der Jahre in listOfYears)
        NrOfY = getEndIndex(listOfYears, end) - getStartIndex(listOfYears, start) + 1
        if NrOfY <1: # Plausibilität
            raise Exception("The Component "+str(d[Comp]["Comp"]["label"]) +
                            " doesnt exist in the chosen Timeframe: "
                            +str(listOfYears)+". Start: "+ str(start)+"; End: "+str(end))

def convertFixCosts(d:dict, listOfYears:list, q:float)-> dict:
    '''
    This combines the Costs from Investment and other annual costs to fixCosts
    and specificCosts, taken by the Framework for Modeling
    It deals with the annuisation of Costs, as well as the Fundings for Investments. THose simply reduce the Invest Costs

    ----------
    d : dict
        dict with ErzeugerData

    listOfYears : list
        list of the years in the calculation
    q : float
        Zinssatz for annuisation

    Returns
    -------
    dict
    '''

    for Comp in d:

        # <editor-fold desc="NrOfY=GetNumberOfYearsInModel">
        # Get the Operation Time of the Component
        start = d[Comp]["NewParam"]["operationStart"]
        end = d[Comp]["NewParam"]["operationEnd"]
        y = end - start  # Anzahl der Jahre, die Die Anlage Existiert

        # Anzahl der Jahre, die Die Anlage im Modell existiert (anzahl der Jahre in listOfYears)
        NrOfY = getEndIndex(listOfYears, end) - getStartIndex(listOfYears, start) + 1
        if NrOfY < 1:  # Plausibilität
            raise Exception("The Component " + str(d[Comp]["Comp"]["label"]) +
                            " doesnt exist in the chosen Timeframe: "
                            + str(listOfYears) + ". Start: " + str(start) + "; End: " + str(end))
        # </editor-fold>

        fixCosts = 0
        specificCosts = 0
        fund=0

        #Funding? 0...1
        for VarName in d[Comp]["NewParam"]:
            if VarName in ["investFund"]:
                fund = d[Comp]["NewParam"][VarName]

        # <editor-fold desc="fixCosts">
        # First: fixCosts, bestehend aus
        #   fixInvestCosts ->       Förderung,         Annuisieren      * Anzahl Jahre
        #   fixCostsPerYear - >                                         * Anzahl Yahre
        for VarName in d[Comp]["NewParam"]:
            if VarName in ["fixInvestCosts"]:
                fixCosts = d[Comp]["NewParam"][VarName]*(1-fund) * get_a(y, q) * NrOfY

        for VarName in d[Comp]["NewParam"]:
            if VarName in ["fixCostsPerYear"]:
                fixCosts = fixCosts + d[Comp]["NewParam"][VarName] * NrOfY
        # </editor-fold>
        d[Comp]["Invest"]["fixCosts"] = fixCosts

        # <editor-fold desc="specificCosts">
        # Second: specificCosts , bestehend aus
        #   spezInvestCosts ->                        Annuisieren     * Anzahl Jahre
        for VarName in d[Comp]["NewParam"]:
            if VarName in ["spezInvestCosts"]:
                specificCosts = d[Comp]["NewParam"][VarName] * get_a(y, q) * NrOfY
        # </editor-fold>
        d[Comp]["Invest"]["specificCosts"] = specificCosts

    return d


def applyOperationFundBEW(d:dict,TotalDict:dict)->dict:
    '''
    This applies the Operational Fundings to HP and ST
    ----------

    Returns
    -------
    dict
    '''
    for comp in d:
        if comp.startswith("WP"):
            if d[comp]["NewParam"]["operationFundBEW"]:
                org=0
                if "costsPerFlowHour" in d[comp]["FlowOut1"]:
                    org=d[comp]["FlowOut1"]["costsPerFlowHour"]
                if d[comp]["Comp"]["COP"]=="Luft":
                    d[comp]["FlowOut1"]["costsPerFlowHour"]=org-TotalDict["TS"]["fundWpLuft"]
                elif d[comp]["Comp"]["COP"]=="Fluss":
                    d[comp]["FlowOut1"]["costsPerFlowHour"]=org-TotalDict["TS"]["fundWpFluss"]
                else:
                    raise Exception("Fund not implemented for COP-Type "+str(d[comp]["Comp"]["COP"]))

        if comp.startswith("ST"):
            if d[comp]["NewParam"]["operationFundBEW"]:
                org=0
                if "costsPerFlowHour" in d[comp]["FlowOut1"]:
                    org=d[comp]["FlowOut1"]["costsPerFlowHour"]
                d[comp]["FlowOut1"]["costsPerFlowHour"]=org-TotalDict["TS"]["fundST"]

    return d
# </editor-fold>

############################################################################################################
# <editor-fold desc="Utility Functions">
def testLengthOfInputdata(InDict,listOfYears):
    le=len(InDict[str(listOfYears[0])].SinkHeat)
    for year in InDict:
        lNew=len(InDict[year].SinkHeat)
        if le!=lNew:
            raise Exception("The length of the Input Data is not consistent!")


# </editor-fold>

# <editor-fold desc="New Preprocessing Modular">

def getEndIndex(listOfYears:list, year:int) ->int:
    '''
    This function returns the index of the last year the Component exists,
    Assumnes that the component goes out of Operation by the start of the "year"
    ----------
    listOfYears : list
        list to be tested
    year : int
        last year of existance of the Component

    Example:
        getEndIndex((2021,2025,2030), 2021)
        Out[37]: 0

        getEndIndex((2021,2025,2030), 2025)
        Out[35]: 1

        getEndIndex((2021,2025,2030), 2026)
        Out[36]: 1

        getEndIndex((2021,2025,2030), 2040)
        Out[39]: 2

    Returns
    -------
    int
    '''
    if year<=listOfYears[0]:
        return -999 # Component is never in operation
        #raise Exception("Out of bounds")
    for ind in listOfYears:
        if ind < year:
            pass
        else:
            return listOfYears.index(ind)-1
    return len(listOfYears)-1 #  Component is in opeation for the whole calculation

def getStartIndex(listOfYears:list, year:int) ->int:
    '''
    This function returns the index of the first year the Component exists

    Example:
        getStartIndex((2021,2025,2030), 2018)
        Out[5]: 0

        getStartIndex((2021,2025,2030), 2021)
        Out[6]: 0

        getStartIndex((2021,2025,2030), 2025)
        Out[7]: 1

        getStartIndex((2021,2025,2030), 2026)
        Out[8]: 2

        getStartIndex((2021,2025,2030), 2040)
        Out[9]: -999
    ----------
    listOfYears : list
        list to be testet
    year : int
        last year of existance of the Component

    Returns
    -------
    int
    '''
    for ind in listOfYears:
        if ind >= year:
            return listOfYears.index(ind)
    return -999
    #raise Exception("Out of bounds")

def existanceTS(startYear:int, EndYear:int, listOfYears:list, VooB:float, orgVal, lenPerYear:int=8760, speicherMaxState=False)->np.ndarray:
    '''
    This functions creates a TS out of min_rel and max_rel, that corresponds to the existance of the component.
    ----------
    startYear : dict
        Start year
    EndYear : int
        End year
    VooB: float
        Value if not existing/out of Bounds
    orgVal: float, np.ndarray
        Value if  existing/in Bounds, can be an array, pd.Series

    Returns
    -------
    dict
    '''

    start=getStartIndex(listOfYears, startYear)                 # 2021      0
    end=getEndIndex(listOfYears, EndYear)                       # 2025      1

    if start==-999 or end==-999:
        ar=np.ones(len(listOfYears) * lenPerYear) * VooB  # create a TS with all values = VooB
    else:
        if isinstance(orgVal,(int,float)):
            ar = np.ones(len(listOfYears) * lenPerYear) * VooB  # create a TS
            ar[start * lenPerYear:(end+1)*lenPerYear] =orgVal
        elif isinstance(orgVal,(pd.Series,pd.DataFrame)):
            if len(orgVal)!=len(listOfYears)*lenPerYear:
                print(len(orgVal))
                print(len(listOfYears)*lenPerYear)
                raise Exception()
            ar=copy.deepcopy(orgVal)
            #ar.iloc[0:start * lenPerYear] = VooB
            mask=ar.index<(start * lenPerYear)
            #ar.iloc[mask] = VooB
            ar.loc[mask] = VooB
            if end + 1 < len(listOfYears):
                ar.iloc[(end + 1) * lenPerYear:len(listOfYears) * lenPerYear] = VooB
            ar=ar.transpose().to_numpy()
        elif isinstance(orgVal,(np.ndarray)):
            ar = copy.deepcopy(orgVal)
            ar[0:start * lenPerYear] = VooB
            if end + 1 < len(listOfYears):
                ar[(end + 1) * lenPerYear:len(listOfYears) * lenPerYear] = VooB
        else:
            raise TypeError

        if speicherMaxState:
            ar=np.append(ar, ar[-1])

    return ar

def yearsOfExistance(operationStart, operationEnd, listOfYears:list)-> int:
    '''
    This functions calculates the number of years the Komponent exists in the calculation.
    To Be  multiplied with the Annuised Invest costs and other fix Costs,
    resulting in the Total Costs over all chosen years
    ----------


    Returns
    -------
    i: int
    '''
    # TODO: Für costsInInvestSizeSegments implementieren
    i=0
    for y in listOfYears:
        if y>= operationStart and y< operationEnd:
            #print(y)
            i=i+1
    return i


# </editor-fold>

# <editor-fold desc="Comps Combined">
def KWKektA(label: str, nominal_val: float, BusFuel: cBus, BusTh: cBus, BusEl: cBus,
            eta_thA: float, eta_elA: float, eta_thB: float, eta_elB: float,min_rel:[float,list], **kwargs)->list:
    '''
    EKT A - Modulation, linear interpolation

    Creates a KWK with a variable rate between electricity and heat production.
    Properties:
        Modulation of Total Power (Fuel) [min_rel, max_rel, nominal value]
        linear interpolation between efficiencies A and B

        Not working: Investment with variable Size

    Parameters
    ----------

    Returns
    -------
        list(cBaseLinearTransformer, cKWK, cKWK)
    '''

    HelperBus = cBus(label='Helper' + label + 'In', media=None)  # balancing node/bus of electricity

    # Transformer 1
    Qin = cFlow(label="Qfu", bus=BusFuel, nominal_val=nominal_val, min_rel=min_rel, **kwargs)
    Qout = cFlow(label="Helper" + label + 'Fu', bus=HelperBus)
    EKTIn = cBaseLinearTransformer(label=label + "In",
                                   inputs=[Qin], outputs=[Qout], factor_Sets=[{Qin: 1, Qout: 1}])
    # EKT A
    EKTA = cKWK(label=label + "A", eta_th=eta_thA, eta_el=eta_elA,
                P_el=cFlow(label="Pel", bus=BusEl),
                Q_fu=cFlow(label="Helper" + label + 'A', bus=HelperBus),
                Q_th=cFlow(label="Qth", bus=BusTh))
    # EKT B
    EKTB = cKWK(label=label + "B", eta_th=eta_thB, eta_el=eta_elB,
                P_el=cFlow(label="Pel", bus=BusEl),
                Q_fu=cFlow(label="Helper" + label + 'B', bus=HelperBus),
                Q_th=cFlow(label="Qth", bus=BusTh))
    return [EKTIn, EKTA, EKTB]

def KWKektB(label: str, BusFuel: cBus, BusTh: cBus, BusEl: cBus,
            segQfu: list[float], segQth: list[float], segPel: list[float], minmax_rel:[int,list]=1, **kwargs)->list:
    '''
    EKT B - On/Off, interpolation with Base Points
    Creates a KWK with a variable rate between electricity and heat production

    Properties:
        On/Off-operation
        Interpolation with Base Points between efficiencies A and B

        Not working:
        Investment with variable Size
        Variation of total Power

    Nominal Value is equal to the max of seqFu

    Parameters
    ----------
    segQfu: list[float]
        Expression with Base Points
        [0, 5, 5, 10]
    segQth: list[float]
        Expression with Base Points
        [0, 3, 3, 9]
    segPel: list[float]
        Expression with Base Points
        [0, 1, 1, 3]

    Returns
    -------
        list(cBaseLinearTransformer, cBaseLinearTransformer, cBaseLinearTransformer)
    '''
    # Testinf for min_rel to only be 0 or 1
    if isinstance(minmax_rel, (float,int)):
        if minmax_rel!=1: raise Exception("min_rel has to be 1, otherwise "+label+" will behave unexpectetly")
    elif all(item == 0 or item == 1 for item in minmax_rel): pass
    else: raise Exception("min_rel must contain only 1 and 0, otherwise "+label+" will behave unexpectetly")

    HelperBus = cBus(label='Helper' + label + 'In', media=None)  # balancing node/bus of electricity

    # Transformer 1
    Qin = cFlow(label="Qfu", bus=BusFuel, nominal_val=max(segQfu), min_rel=minmax_rel, max_rel=minmax_rel, **kwargs)
    Qout = cFlow(label="Helper" + label + 'Fu', bus=HelperBus)
    EKTIn = cBaseLinearTransformer(label=label + "In",
                                   inputs=[Qin], outputs=[Qout], factor_Sets=[{Qin: 1, Qout: 1}])

    # Transformer Strom
    P_el = cFlow(label="Pel", bus=BusEl)
    Q_fu = cFlow(label="Helper" + label + 'A', bus=HelperBus)
    segs = {Q_fu: segQfu.copy(), P_el: segPel.copy()}
    EKTA = cBaseLinearTransformer(label=label + "A", outputs=[P_el], inputs=[Q_fu], segmentsOfFlows=segs)

    # Transformer Wärme
    Q_th = cFlow(label="Qth", bus=BusTh)
    Q_fu = cFlow(label="Helper" + label + 'B', bus=HelperBus)
    segs = {Q_fu: segQfu.copy(), Q_th: segQth.copy()}
    EKTB = cBaseLinearTransformer(label=label + "B", outputs=[Q_th], inputs=[Q_fu], segmentsOfFlows=segs)

    return [EKTIn, EKTA, EKTB]
# </editor-fold>