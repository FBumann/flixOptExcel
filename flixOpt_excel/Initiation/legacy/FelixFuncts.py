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

from flixBasicsPublic import *
from flixComps import *
from flixPlotHelperFcts import *
from flixStructure import *
import flixPostprocessing as flixPost
from flixPlotHelperFcts import *
import pickle

# all utility FUnctions

def merge_dict(dict1, dict2):
    # https: // stackoverflow.com / questions / 43797333 / how - to - merge - two - nested - dict - in -python
    for key, val in dict1.items():
        if type(val) == dict:
            if key in dict2 and type(dict2[key] == dict):
                merge_dict(dict1[key], dict2[key])
        else:
            if key in dict2:
                dict1[key] = dict2[key]

    for key, val in dict2.items():
        if not key in dict1:
            dict1[key] = val

    return dict1

####### Functions for Dict nesting from Dataframe
def nest(d: dict) -> dict:
    # Source. https://stackoverflow.com/questions/50929768/pandas-multiindex-more-than-2-levels-dataframe-to-nested-dict-json
    result = {}
    for key, value in d.items():
        target = result
        for k in key[:-1]:  # traverse all keys but the last
            target = target.setdefault(k, {})
        target[key[-1]] = value
    return result

def df_to_nested_dict(df: pd.DataFrame) -> dict:
    d = df.to_dict()
    return {k: nest(v) for k, v in d.items()}


# load json module
def saveToPickle(dict:dict, path:str):
    import pickle
    # create a binary pickle file
    f = open("file.pkl", "wb")

    # write the python object (dict) to pickle file
    pickle.dump(dict, f)

    # close file
    f.close()


def filterLeapYear(TS): # TODO:This Func doesnt really work, ist just saved for later completion and useage
    # filter out all timesteps of the 29th february

    filter = np.array([], dtype=bool)
    for dt in TS:
        if ((dt.month == 2) & (dt.day == 29)):
            filter = np.append(filter, False)
        else:
            filter = np.append(filter, True)

    aTimeSeries = TS[filter]
    return aTimeSeries

def getDepthOfDict(d:dict):
    from collections import deque

    queue = deque([(id(d), d, 1)])
    memo = set()
    while queue:
        id_, o, level = queue.popleft()
        if id_ in memo:
            continue
        memo.add(id_)
        if isinstance(o, dict):
            queue += ((id(v), v, level + 1) for v in o.values())
    return level
