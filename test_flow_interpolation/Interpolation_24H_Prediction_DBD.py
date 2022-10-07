import numpy as np
import pandas as pd
from icecream import ic
import os 

#ic.configureOutput(includeContext=True)

#ic(repr('D:\\...\\Miri\\exp.xlsx'))
#data = pd.read_excel (os.getcwd()+'\\exp.xlsx', engine='openpyxl')
df = pd.read_excel ('total_temp.xlsx', engine='openpyxl')
df.columns = ['DATE', 'VALUE']
#df = pd.DataFrame(data, columns= ['DATE','VALUE'])

daily_values = df['VALUE'].to_numpy()
length_daily_values = len(daily_values)
ic(daily_values)

daily_values_array = []
daily_values_label = []

for i in range(24,length_daily_values,25):

    if np.isnan(daily_values[i]):
        continue
    
    if np.isnan(daily_values[i:(i+24)]).any():
        continue
    daily_values_label.append(round(i/24))
    daily_values_array.append(daily_values[i:(i+24)])

ic(daily_values_array)
ic(daily_values_label)

interpolated_daily_values = []

#ic(daily_values_label)
for i in range(0,len(daily_values_label)-1):
    
    interpolated_daily_values.append(daily_values_array[i])
    delta = daily_values_label[i+1] - daily_values_label[i]
    #ic(delta)
    for j in range(1,delta):
        #ic(j)
        left_interpolation = daily_values_array[i]*(delta-j)
        right_interpolation = daily_values_array[i+1]*(j)
        interpolated_value = np.add(left_interpolation,right_interpolation)
        interpolated_value /=delta
        interpolated_daily_values.append(interpolated_value)

ic(len(interpolated_daily_values))
ic(daily_values[32:33])

for i in range(1,280):
    index = i*24 -1
    daily_values[index:(index+24)] = interpolated_daily_values[i-1]

df.to_csv('total_interpolated.csv',index=False)