import numpy as np

import pandas as pd

dti = pd.date_range("2018-01-01", periods=8760, freq="H")

totaldata = pd.read_excel('IMPEXP_AB_SK_2030.xlsx', sheet_name='Total', usecols='B')
totaldata = totaldata.replace({'0':np.nan, 0:np.nan})
totaldata.index = dti

#importdata = pd.read_excel('IMPEXP_NS_2030.xlsx', sheet_name='Imp', usecols='E')
#importdata = importdata.replace({'0':np.nan, 0:np.nan})
#importdata.index = dti


#exp_data_int = totaldata.interpolate(method='time')
#imp_data_int = importdata.interpolate(method='time')

#importdata.to_excel('imp_temp.xlsx')
totaldata.to_excel('total_temp.xlsx')
#exp_data_int = exp_data_int.replace({np.nan:0})
#imp_data_int = imp_data_int.replace({np.nan:0})

#exp_data_int.to_excel('inter_exp.xlsx')
#imp_data_int.to_excel('inter_imp.xlsx')
