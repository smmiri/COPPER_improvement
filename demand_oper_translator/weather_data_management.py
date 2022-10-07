import glob
import os
import pandas as pd

cwd = os.getcwd()
pwd = os.path.dirname(cwd)
dir = input('Insert the folder containing case data:\n')
os.chdir(f'{pwd}/{dir}/weather_data')

extension = 'csv'
st1_name = [i for i in glob.glob('en_climate_hourly_AB_3012206*.{}'.format(extension))]
st1_temp = pd.concat([pd.read_csv(f, usecols=['Temp (°C)']) for f in st1_name], axis=0)
st1_temp.columns = ['st1_temp']

st2_name = [i for i in glob.glob('en_climate_hourly_AB_3031094*.{}'.format(extension))]
st2_temp = pd.concat([pd.read_csv(f, usecols=['Temp (°C)']) for f in st2_name], axis=0)
st2_temp.columns = ['st2_temp']


temp = pd.concat([st1_temp, st2_temp], axis=1)
temp['mean'] = temp.mean(axis=1)
temp.index = pd.to_datetime(temp.index, dayfirst=True, unit='h', origin='2010-01-01')
temp.to_csv('temps.csv')
