#!/usr/bin/env python
# coding: utf-8

# In[62]:


import glob
import os
import pandas as pd

def fixpath(path):
    if path.startswith("C:"): return "/mnt/c/" + path.replace("\\", "/")[3:]
    else:
        pass
    return path

os.chdir(fixpath(r'C:\Users\smoha\Downloads\SILVER_BC_Cascade\SILVER_BC_Cascade\SILVER_Data\user_inputs\Hydro_Data-BC_Cascade'))

# Read the hydro data
cascade_cf = pd.read_csv('hydro_cascade.csv', index_col=0)
src_index = cascade_cf.index
cascade_cf.reset_index(inplace=True)
cascade_cf.index = pd.to_datetime(cascade_cf.index, unit='h', dayfirst=True, origin='2050-01-01')

cascade_pmin = pd.read_csv('hydro_cascade_pmin.csv', index_col=0)
cascade_pmin.reset_index(inplace=True)
cascade_pmin.index = pd.to_datetime(cascade_pmin.index, unit='h', dayfirst=True, origin='2050-01-01')

hourly_cf = pd.read_csv('hydro_hourly.csv', index_col=0)
hourly_cf.reset_index(inplace=True)
hourly_cf.index = pd.to_datetime(hourly_cf.index, unit='h', dayfirst=True, origin='2050-01-01')

hourly_pmin = pd.read_csv('hydro_hourly_pmin.csv', index_col=0)
hourly_pmin.reset_index(inplace=True)
hourly_pmin.index = pd.to_datetime(hourly_pmin.index, unit='h', dayfirst=True, origin='2050-01-01')


# In[33]:


# Select the periods of interest

periods = input("Enter the period of interest in dates (e.g. 2050-01-01 to 2050-03-01): ")
periods = periods.split('to')
periods = [pd.to_datetime(periods[0]), pd.to_datetime(periods[1])]
timeline = pd.date_range(periods[0], periods[1], freq='H')
print(timeline)


# In[45]:


# Selecting the cascades of interest

print("The following cascades are available: ")
print(cascade_cf.columns[1:-3])
imp_cascades = input("Enter the cascade(s) of interest from the list (e.g. 1,2,3): ")
imp_cascades = imp_cascades.split(',')
imp_cascades = [int(i) for i in imp_cascades]
imp_cascades = [cascade_cf.columns[i] for i in imp_cascades]
print(imp_cascades)


# In[63]:


# Changing the minimum capacity of the cascades

imp_factor = input("Enter the factor by which the cascading minimum caoacity should be increased (e.g. 1.2): ")
imp_factor = float(imp_factor)
imp_cascade_pmin = pd.DataFrame(cascade_pmin)
imp_cascade_pmin.loc[(cascade_pmin.index >= timeline[0]) & (cascade_pmin.index <= timeline[-1]), imp_cascades] = cascade_pmin.loc[(cascade_pmin.index >= timeline[0]) & (cascade_pmin.index <= timeline[-1]), imp_cascades] * imp_factor


# In[59]:


# Writing to new files

scen = input("Enter the scenario name: ")
imp_cascade_pmin.to_csv('hydro_cascade_pmin_' + scen + '.csv', index=False)

