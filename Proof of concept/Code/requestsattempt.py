# -*- coding: utf-8 -*-
"""
Created on Fri Oct 13 09:41:07 2023

@author: hanif.jetha
"""
import requests
from requests_ntlm import HttpNtlmAuth
import getpass
import pandas as pd
import numpy as np

username = "hanif.jetha@ofwat.gov.uk"

# "hanif.jetha@ofwat.gov.uk"

url = "https://fountain01/Fountain/rest-services/report/flattable/20645"

data = requests.get(url, verify = False, auth= HttpNtlmAuth(username, getpass.getpass()), timeout=20).text

#df = pd.read_xml(data)
df = pd.read_xml(data, dtype = str)
df.drop(columns = ['cell_0'],inplace = True)
df.replace(np.nan, None, inplace = True)
df.columns = df.iloc[0]
df = df[1:]
print(df)