import math
import os
import pandas as pd
from icecream import ic
import csv

with open('windcf.csv') as csv_file:
    reader = csv.reader(csv_file)
    windcf = dict(reader)

for k in windcf:
    windcf[k]=float(windcf[k])

with open('solarcf.csv') as csv_file:
    reader = csv.reader(csv_file)
    solarcf = dict(reader)

for k in solarcf:
    solarcf[k]=float(solarcf[k])

coordinate_df = pd.read_excel("coordinate.xlsx", sheet_name="coordinate_system")
coordinate_df = coordinate_df.dropna(how='all')
sample_df = pd.read_csv('Wind_Generation_Data_285-105_2018.csv', index_col=0)
ic(sample_df)
#windcf_df = pd.read_csv('windcf.csv', header=None)
#os.chdir(os.getcwd()+'/temp')
os.chdir(os.getcwd()+'/unix_temp')
dir = os.getcwd()
for i, row in coordinate_df.loc[coordinate_df['grid cell'] != ''].iterrows():
    latitude = row['lat']
    latitude_merra = int(round((row['lat']+90)*2,0))
    longitude = row['lon']
    longitude_merra = int(round((row['lon']+180)*1.5,0))
    grid_cell = int(row['grid cell'])
    wind_list = []
    for j in range(1, 8761):
        wind_list.append(windcf[str(j) + '.' + str(grid_cell)])
    wind_list_df = pd.DataFrame(wind_list, index=sample_df.index, columns=sample_df.columns)
    new_dir = os.getcwd() + '/Wind_Generation_Data/' + str(latitude_merra) + '-' + str(longitude_merra)
    os.makedirs(new_dir)
    os.chdir(new_dir)
    wind_list_df.to_csv('Wind_Generation_Data_' + str(latitude_merra) + '-' + str(longitude_merra) + '_2018.csv')
    del wind_list
    del wind_list_df
    os.chdir(dir)

    solar_list = []
    for j in range(1, 8761):
        solar_list.append(solarcf[str(j) + '.' + str(grid_cell)])
    solar_list_df = pd.DataFrame(solar_list, index=sample_df.index, columns=sample_df.columns)
    new_dir = os.getcwd() +'/Solar_Generation_Data/'+str(latitude_merra)+'-'+str(longitude_merra)
    os.makedirs(new_dir)
    os.chdir(new_dir)
    solar_list_df.to_csv('Solar_Generation_Data_' + str(latitude_merra) + '-' + str(longitude_merra) + '_2018.csv')
    del solar_list
    del solar_list_df
    os.chdir(dir)



#os.chdir(os.getcwd()+'/temp')

'''
for i in range(1,2279):
    wind_list = []
    for j in range(1,8761):
        wind_list.append(windcf[str(j)+'.'+str(i)])
    pd.DataFrame(wind_list).to_csv('wind_' + str(i) + '.csv', header=None)
    del wind_list

for i in range(1,2279):
    solar_list = []
    for j in range(1,8761):
        solar_list.append(solarcf[str(j)+'.'+str(i)])
    pd.DataFrame(solar_list).to_csv('solar_' + str(i) + '.csv', header=None)
    del solar_list

'''