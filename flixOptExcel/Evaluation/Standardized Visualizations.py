import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from copy import deepcopy
from typing import Literal

from flixOptExcel.Evaluation.flixPostprocessingXL import flixPostXL
from flixOptExcel.Evaluation.HelperFcts_post import resample_data

class _cGraficData(dict):
    '''
    This class serves as a storage container for a bunch of data for easy evaluation.
    It is the Mother class for several data computation classes.
    Its main purpose is to store the data in a standardized dict structure. For this the _update() function is used
    '''
    def __init__(self, calc1 :flixPostXL):
        super().__init__()
        self.calc1 = deepcopy(calc1) # TODO: Is this necessary? Performance issues?

    def _update(self, name :str, y_unit :str, data :pd.DataFrame, color_map :dict =None):
        # Used to get a standardized structure
        self['infos'] = {}
        self['infos']["name"] =name
        self['infos']["y_unit"] = y_unit
        self['data'] = data
        self['infos']["color_map"] = color_map

    # TODO: Create a function that plots the data and saves it in a custom location.
    def plot(self, sample_rate: Literal["Y", "D", "H"], path_to_save: str = None,
             # style:Literal['line', 'bar', 'barh'],
             # type:Literal['HTML','png']

             ):
        df = self["data"][sample_rate]
        fig = px.bar(data_frame=df,
                     title=self["infos"]["name"],
                     # y = df.columns,
                     color_discrete_map=self["infos"]["color_map"])
        fig.update_layout(yaxis_title=self["infos"]["y_unit"])

        # Create a figure with stacked bars
        fig = px.bar(df, barmode='stack',
                     title=self["infos"]["name"], labels={'value': self["infos"]["y_unit"]})

        # Add a line trace
        fig.add_trace(go.Scatter(x=df['Category'], y=df['LineValue'], mode='lines+markers', name='LineValue'))

        # Update layout
        fig.update_layout(title='Combined Figure with Stacked Bars and Line')

        fig.show()
        # fig.write_image(path_to_save) # TODO Continue here 28.12
        # fig.write_html()


class cGraficDataExcel(_cGraficData):
    def __init__(self, fill_function_name, calc: flixPostXL):
        '''
        This class serves as a storage container for a bunch of data for easy evaluation
        :param fill_function_name: Choose the function to calculate the data.
        Possible functions: 'Fernwärme Last und Verluste', 'Installierte Leistung'
        :param calc: the calculation results
        '''
        super().__init__(calc)

        self.basic_colors = px.colors.qualitative.Light24

        # Mapping of fill function names to actual functions
        fill_functions = {
            'Fernwärme Last und Verluste': self._fernwaerme_last_und_verluste,
            'Installierte Leistung': self._invest_values
            # 'custom_fill2': self._custom_fill_function2,
            # Add more fill functions as needed
        }

        # Call the selected fill function if one is provided
        fill_function = fill_functions.get(fill_function_name)
        if fill_function:
            fill_function()
        else:
            raise Exception(f"This is no valid fill function_name. Choose from {fill_functions.keys()}")

    # Costum functions to create a dataset
    def _fernwaerme_last_und_verluste(self) -> dict:
        # caluculation of the data
        df_fernwaerme_last = self.calc1.to_dataFrame("Fernwaerme", "out", grouped=True).filter(like='Waermelast')
        df_fernwaerme_last = self.calc1.reorder_columns(df_fernwaerme_last)

        # Extract the relevant part of the dictionary for the specified keys
        color_map = {key: self.calc1.color_map[key] for key in df_fernwaerme_last.columns if
                     key in self.calc1.color_map}
        style_map = {key: "bar" for key in df_fernwaerme_last.columns}

        data = {}
        # Resampling of the data
        for sample_rate in ("H", "D", "Y"):
            df_summed = resample_data(df_fernwaerme_last, self.calc1.years, sample_rate, "mean")
            df_verluste_summed = (df_summed['Waermelast_Netzverluste_Qth'] / df_summed.sum(axis=1) * 100).rename(
                "Verlust[%]")
            data[sample_rate] = pd.concat([df_summed, df_verluste_summed], axis=1)

        color_map["Verlust[%]"] = self.basic_colors[0]
        style_map["Verlust[%]"] = "line"
        # information about the data
        self._update(name="Fernwärme Last und Verluste", y_unit="MW", data=data, color_map=color_map,
                     style_map=style_map)
        return self

    def _invest_values(self) -> dict:
        # caluculation of the data
        val = self.calc1.get_invest_results(flows=True, storages=False)
        ex = self.calc1.get_exist_values()

        product_dict = {}
        # Iterate through keys that are common to both dictionaries
        for key in set(val.keys()) & set(ex.keys()):
            value1 = val[key]
            value2 = ex[key]
            # Assuming all arrays are one-dimensional
            if isinstance(value1, float) and isinstance(value2, np.ndarray):
                # Element-wise multiplication for arrays using NumPy
                product_dict[key] = value1 * value2

        df = pd.DataFrame(product_dict)

        # Resampling of the data, if possible in a loop
        data = {}
        for sample_rate in ("H", "D", "Y"):
            data[sample_rate] = resample_data(df, self.calc1.years, sample_rate, "mean")

        # information about the data
        self._update(name="Installierte Leistung", y_unit="MW", data=data)
        return self

        # Grafik typ 1: Bus-Bilanz - Fernwärmeerz, Stromerzeugung
        ##self.to_dataFrame("Fernwaerme")

        # Grafik Typ 2: Investment Results (Leistung/Kapa)

        # Grafik Typ 3: Effekte (Kosten/CO2/...)

        # Grafik Typ 4: Sources (Energieträger ???

        # Grafik Typ 4: Speicher Füllstand

    def _invest_costs(self) -> dict:
        pass

