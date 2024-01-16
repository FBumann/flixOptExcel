# -*- coding: utf-8 -*-
"""
Created on Thu Jun 16 11:19:17 2022
developed by Felix Panitz* and Peter Stange*
* at Chair of Building Energy Systems and Heat Supply, Technische Universität Dresden
"""
import numpy as np
import datetime
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import openpyxl
import pprint

from flixBasicsPublic import *
from flixComps import *
from flixPlotHelperFcts import *
from flixStructure import *
import flixPostprocessing as flixPost
from flixPlotHelperFcts import *
import copy

def dictToDF(mydict, names):
    '''
        Create a Dataframe from a Dict f.ex. for Operation Costs of the Components

         Parameters
         ----------
         mydict : dict
             Results dict
         Comps : list
            List of Names of the Components to be looked at

         Returns
         -------
         pd.Dataframe
        '''

    df=pd.DataFrame()
    #Iterate over every Component
    for CompName in names:
        #Create a df for every Component
        df1 = pd.DataFrame(index=[CompName])
        for key in mydict:
            if key.startswith(CompName):
                df2 = pd.DataFrame(data=[mydict[key]], index=[CompName], columns=[str(key).split("_", 2)[-1]])
                df1= pd.concat([df1, df2], axis=1)
        #join df of Component to dataFrame
        df= pd.concat([df, df1])
    return df

def writeToExcelOutput(data, wsName, path='/Users/felix/Documents/Uni/Diplomarbeit/Energiesystemmodell/'
                                          'flixOpt/flixOpt-main/Own Modells/proj_fin/resources/resources.xlsx'):
    '''
    Writes a DataFrame to an existing Worksheet in a excelFile

    Parameters
    ----------
    data : pd.Dataframe
        Data to be written
    path : str
        String of current file to write into
    wsName : str
        Name of the worksheet in the excel file

    Returns
    -------

    '''

    # TODO: WATCH OUT: Clearing of old data only from A1:K10.000 - change if necessary (RunTime!)
    wb = openpyxl.load_workbook(path)
    #saving the old output as a copy
    wb.save(str(path.split('.')[0])+'_old.xlsx')
    # initialize stuff
    writer = pd.ExcelWriter(path, engine='openpyxl')
    writer.book = wb
    ws=wb[wsName]
    #clear old data
    for row in ws['A1:AA50000']:
      for cell in row:
        cell.value = None

    #write new data
    writer.sheets = dict((ws.title, ws) for ws in wb.worksheets)
    #write it all out at the first row in the excel sheet
    data.to_excel(writer, sheet_name = wsName, startrow =0,index=True)

    wb.save(path)
    wb.close()

def rs(df:pd.DataFrame, unit:str, Auflösung:str=None)-> pd.DataFrame:
    # Resampling of the Data Frame to other timeFrame, this expects to have a Date and Time in the index
    # and furthermore change the Unit of the Dataframe
    # this function expects the input to be in kWh or €
    if not Auflösung==None:
        df=df.resample(Auflösung).sum()

    if unit not in ["€", "t€", "Mio.€", "kWh", "MWh", "GWh"]:
        raise Exception("unit has to be one of the following: [€,t€,Mio.€,kWh,MWh,GWh]")
    if unit == "MWh":
        df = df.divide(1000)
    if unit == "GWh":
        df = df.divide(1000000)
    if unit == "t€":
        df = df.divide(1000)
    if unit == "Mio.€":
        df = df.divide(1000000)
    return df

# FUnction to create a stacked bar Plot
def stackedBarPlot(df:pd.DataFrame,title:str, unit:str,addValues=True, Auflösung=None):
    # https://stackoverflow.com/questions/41296313/stacked-bar-chart-with-centered-labels
    plt.style.use('ggplot')
    df=df.dropna(axis=1, how='all') #drop all nan columns

    # Resampling of the Data Frame, this expects to have a Date and Time in the index
    if Auflösung!=None: df=df.resample(Auflösung).sum()

    # <editor-fold desc="Unit conversion">
    if unit not in["€","t€","Mio.€","kWh","MWh","GWh"]:
        raise Exception("unit has to be one of the following: [€,t€,Mio.€,kWh,MWh,GWh]")

    if unit=="MWh":
        df=df.divide(1000)
    if unit == "GWh":
        df=df.divide(1000000)
    if unit=="t€":
        df=df.divide(1000)
    if unit == "Mio.€":
        df=df.divide(1000000)
    # </editor-fold>

    ax=df.plot(kind='bar', stacked=True, figsize=(8, 6), rot=0)
    plt.ylabel(unit)
    plt.title(label=title)
    plt.grid(visible=True, axis="y")

    # .patches is everything inside of the chart
    if addValues:
        for rect in ax.patches:
            # Find where everything is located
            height = rect.get_height()
            width = rect.get_width()
            x = rect.get_x()
            y = rect.get_y()

            # The height of the bar is the data value and can be used as the label
            label_text = f'{height:.0f}'  # f'{height:.2f}' to format decimal values

            # ax.text(x, y, text)
            label_x = x + width / 2
            label_y = y + height / 2

            # plot only when height is greater than specified value
            if height > 0:
                ax.text(label_x, label_y, label_text, ha='center', va='center', fontsize=12)

    ax.set_xlabel('Total Sum: ' + str(round(df.sum(axis=1).sum(axis=0)))+' '+unit)
    return ax

def stackedBarPlotFromResultsDF(df:pd.DataFrame,title:str, unit:str,addValues=True, Auflösung=None):
    # https://stackoverflow.com/questions/41296313/stacked-bar-chart-with-centered-labels
    plt.style.use('ggplot')
    df=df.dropna(axis=1, how='all') #drop all nan columns

    # Resampling of the Data Frame, this expects to have a Date and Time in the index
    if Auflösung!=None: df=df.resample(Auflösung).sum()

    # <editor-fold desc="Unit conversion">
    if unit not in["€","t€","Mio.€","kWh","MWh","GWh"]:
        raise Exception("unit has to be one of the following: [€,t€,Mio.€,kWh,MWh,GWh]")

    if unit=="MWh":
        df=df.divide(1000)
    if unit == "GWh":
        df=df.divide(1000000)
    if unit=="t€":
        df=df.divide(1000)
    if unit == "Mio.€":
        df=df.divide(1000000)
    # </editor-fold>

    ax=df.plot(kind='bar', stacked=True, figsize=(8, 6), rot=0)
    plt.ylabel(unit)
    plt.title(label=title)
    plt.grid(visible=True, axis="y")

    # .patches is everything inside of the chart
    if addValues:
        for rect in ax.patches:
            # Find where everything is located
            height = rect.get_height()
            width = rect.get_width()
            x = rect.get_x()
            y = rect.get_y()

            # The height of the bar is the data value and can be used as the label
            label_text = f'{height:.0f}'  # f'{height:.2f}' to format decimal values

            # ax.text(x, y, text)
            label_x = x + width / 2
            label_y = y + height / 2

            # plot only when height is greater than specified value
            if height > 0:
                ax.text(label_x, label_y, label_text, ha='center', va='center', fontsize=12)

    ax.set_xlabel('Total Sum: ' + str(round(df.sum(axis=1, numeric_only=True).sum(axis=0)))+' '+unit)
    return ax

def dfFromNestedDict(dict:dict,aimLength:int)->pd.DataFrame:

    '''
    Writes a DataFrame from a nested dict and converts values to match the length of the DataFrame.
    If the length of the data in the dict is >1, Only takes numeric data

    Parameters
    ----------
    dict : dict
        nested dict
    aimLength : int
        length of the data / Nr. of Timesteps
        use: len(calc1.results_struct.time.timeSeries)

    Returns
    -------
    '''

    df_melt = pd.json_normalize(dict, sep='>>').melt()
    df_final = df_melt['variable'].str.split('>>', expand=True)
    df_final.columns = [f'Label_{name+1}' for name in df_final.columns]

    # Extra: Setzte Column names, bestehend aus den Einzelnamen der Keys

    colList = df_melt['variable'].str.split('>>').to_list()
    columnList = list()
    for subList in colList:
        #subList.pop(0)
        columnList.append("_".join(subList))

    df_final = df_final.set_axis(columnList, axis=0)
    df_melt = df_melt.set_axis(columnList, axis=0)

    # create a Dataframe from the results dict values
    dfData = pd.DataFrame()
    i = -1
    for v in df_melt.value:
        i = i + 1
        vLe = v
        if type(vLe) == np.ndarray:
            if len(vLe) == 1:
                vLe = np.ones(aimLength) * vLe.values
            elif len(vLe) == aimLength:
                pass
            elif len(vLe) == aimLength + 1:  # This exists for the timesteps with an extra step at the end. Has to be eliminated #TODO: Why??
                vLe=vLe[:-1]
                #vLe = np.empty(aimLength)
                #vLe[:] = np.nan
            else:
                vLe = vLe[:aimLength]
                pass #TODO: Change
                #raise Exception("Not viable lenght: "+str(len(vLe)))

        elif isinstance(vLe, (float, int)):
            vLe = np.ones(aimLength) * vLe
        elif vLe == None:
            vLe = np.empty(aimLength)
            vLe[:] = np.nan
        elif isinstance(vLe, (str)):
            vLe=vLe*aimLength
        else:
            raise Exception("Not implemented for Datatype: "+str(type(vLe)))
        df1 = pd.DataFrame(vLe, index=np.arange(stop=aimLength), columns=[columnList[i]]).T
        dfData = pd.concat([dfData, df1])

    ResultsDF = pd.concat([df_final, dfData], axis=1)

    # ResultsDF = ResultsDF.drop(["Label_0"], axis=1)

    return ResultsDF.T

class cOutputDF(pd.DataFrame):
    """
    Klasse cOutputDF
    This class is made for better handling of the resources Dataframe
    """
    def __init__(self,data, **kwargs):
        super().__init__( data,**kwargs)
        # create a Timeseries that matches the length of the Dataframe (hourly increments)
        TimeSeries=np.ndarray([])
        if False:
            for year in listOfYears:
                TS = datetime.datetime(year, 1, 1) + np.arange(8760) * datetime.timedelta(hours=1)
                if TimeSeries.shape==():
                    TimeSeries=TS.astype('datetime64')
                else:
                    TimeSeries=np.concatenate((TimeSeries,TS.astype('datetime64')))


    def dropLabels(self):
        '''
        Drop the labels from the Dataframe.
        Returns the Dataframe containing only Data
        Returns a instance of cOutputDF

        Parameters
        ----------

        Returns
        -------
        cOutputDF
        '''
        labels=list()
        for index in self.index:
            if type(index)==str:
                labels.append(index.format())
        return cOutputDF(self.drop(labels))

    def getLabelsDF(self):
        '''
        Returns the Dataframe containing only the Labels
        Returns a instance of cOutputDF

        Parameters
        ----------

        Returns
        -------
        cOutputDF
        '''
        labels=list()
        for index in self.index:
            if type(index)==str:
                labels.append(index.format())

        # filter the DataFrame based on the list of index names
        return cOutputDF(self.loc[labels])

    def filterDF(self,Label_1:str=None,Label_2:str=None,Label_3:str=None, Label_4:str=None,startsWith:bool=False):
        df=self
        if not startsWith:
            if Label_1 !=None:
                df=df.loc[:, df.loc["Label_1"] == Label_1]
            if Label_2 != None:
                df=df.loc[:, df.loc["Label_2"] == Label_2]
            if Label_3 != None:
                df=df.loc[:, df.loc["Label_3"] == Label_3]
            if Label_4 != None:
                df = df.loc[:, df.loc["Label_4"] == Label_4]
        else:
            df=copy.deepcopy(df).fillna("Platzhalter")
            if Label_1 !=None:
                df=df.loc[:, df.loc["Label_1"].str.startswith(Label_1)]
            if Label_2 != None:
                df=df.loc[:, df.loc["Label_2"].str.startswith(Label_2)]
            if Label_3 != None:
                df=df.loc[:, df.loc["Label_3"].str.startswith(Label_3)]
            if Label_4 != None:
                df=df.loc[:, df.loc["Label_4"].str.startswith(Label_4)]
        return cOutputDF(df)


    def groupByLabel(self, FullLabel:bool=True, LabelNr:int=2, NrOfLetters:int=2):
        '''
        Groups the Dataframe by the value stored in one of the labels (Index 1-5)
        if FullLabel=False, group only ba fraction of the Label (f.ex. only first 2 letters)
        Returns a instance of cOutputDF
        TODO: Dangerous!!. Maybe always use 3 Letters in InputSheet (WP-> WaP, ST->SoT

        Parameters
        ----------
        NrOfLetters : int
            Number of letters to group by
        LabelNr : int
            Number of the Label to group by


        Parameters
        ----------

        Returns
        -------
        cOutputDF
        '''
        groups = {}
        df=self
        df=df.fillna(value="None") # For handling None Labels

        for i in range(len(df.columns)):
            lab=df.loc["Label_"+str(LabelNr)][i]
            if FullLabel:
                prefix =lab
            else:
                if type(lab)==str:
                    prefix = lab[:NrOfLetters]
                else:
                    prefix = str(lab)

            col=df.columns[i] # column name
            if prefix not in groups:
                groups[prefix] = [col]
            else:
                groups[prefix].append(col)
        dfNew = pd.DataFrame()
        for prefix in groups:
            # filter dataframe based on list of columns
            filtered_df = df[groups[prefix]]

            # sum up values per row (axis=1)
            dfNew[prefix] = filtered_df.sum(axis=1)

        return cOutputDF(dfNew)

    # Still a valid function 23.11.2023
    def rs(self, loY:list, Auflösung: str = "D",type:str="mean",lengthPerY:int=8760):
        '''
        Returns the Dataframe with resampled Data
        # TODO: CHeck if is always working (GAP YEARS!!) Maybe change to other years as indexes
        # TODO Check if ist possible and viable to use the real Years and leave the index as Datetimes

        Parameters
        ----------
        type: str
            choose from: min, max, mean, sum
        Returns
        -------
        cOutputDF
        '''
        df=self.dropLabels()
        #index = pd.date_range('1/1/2000', periods=df.shape[0], freq='H')
        #df = df.set_index(index)
        df.index = range(len(df))

        i=0
        if lengthPerY==8760:
            freq='H'
        elif lengthPerY==365:
            freq='D'
        else: raise Exception()

        for y in loY:
            dt=pd.date_range(start='1/1/'+str(y), end='01/01/'+str(y+1), freq=freq)[:-1]
            dt = dt[~((dt.month == 2) & (dt.day == 29))] # Schaltjahr!!
            df.loc[i*lengthPerY:(i+1)*lengthPerY-1, 'Timestamp'] = dt
            i=i+1
        df = df.set_index('Timestamp')

        if type=="sum":
            df = df.resample(Auflösung).sum()
        elif type=="mean":
            df = df.resample(Auflösung).mean()
        elif type=="min":
            df = df.resample(Auflösung).min()
        elif type=="max":
            df = df.resample(Auflösung).max()
        else:
            raise Exception()

        # TODO: Drop all rows, which arent in the Years specified in loY
        lst=list()
        for row in df.index:
            if row.year not in loY:
                lst.append(row)
        df=df.drop(index=lst)

        df = df.loc[~((df.index.month == 2) & (df.index.day == 29))]  # Schaltjahr!!

        if Auflösung=="Y":
            df = df.set_index(df.index.year)


        return cOutputDF(pd.concat([self.getLabelsDF(), df]))

    def splitByYear(self):

        df=self.dropLabels()
        lenYear = 8760
        if len(df) % lenYear != 0:  # Rest?
            raise Exception("Error")

        NrY = int(len(df) / lenYear)
        yearlyDF = np.array_split(df, NrY)
        lst=list()
        for i in range(len(yearlyDF)):
            lst.append(cOutputDF(pd.concat([self.getLabelsDF(), yearlyDF[i]], axis=0)))

        return lst

    def to_excelDesktop(self, filename: str = "Test.xlsx", sheetname: str = "Tabelle 1", index: bool = True):
        print("########## write excel - To Desktop ############")
        import os
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')

        if os.name == 'nt':  # 'nt' represents Windows
            directory_separator = '\\'
        else:
            directory_separator = '/'

        path = desktop_path + directory_separator + filename
        self.to_excel(path, sheet_name=sheetname, index=index)

    def sortBy(self,sortByLabel:str):
        df=self.dropLabels()
        df=df.sort_values(sortByLabel)

        return cOutputDF(pd.concat([self.getLabelsDF(), df]))




def group_columns_by_Comp(df):
    groups = {}
    for col in df.columns:
        prefix = col.split('_')[0]
        if prefix not in groups:
            groups[prefix] = [col]
        else:
            groups[prefix].append(col)
    dfNew=pd.DataFrame()
    for prefix in groups:
        # filter dataframe based on list of columns
        filtered_df = df[groups[prefix]]

        # sum up values per row (axis=1)
        dfNew[prefix] = filtered_df.sum(axis=1)

    return dfNew

def group_columns_by_firstTwoLetters(df):
    groups = {}
    for col in df.columns:
        prefix = col[:2]
        if prefix not in groups:
            groups[prefix] = [col]
        else:
            groups[prefix].append(col)
    dfNew=pd.DataFrame()
    for prefix in groups:
        # filter dataframe based on list of columns
        filtered_df = df[groups[prefix]]

        # sum up values per row (axis=1)
        dfNew[prefix] = filtered_df.sum(axis=1)

    return dfNew



