from glob import glob
from openpyxl import Workbook, load_workbook
import os
import pandas as pd
import shutil
import time

cwd = os.getcwd()
pwd = os.path.dirname(cwd)

naming = input("Input the iteration number province, and scenario in the correct format:\n(iteration_province_scenario in the format, iteration_elecscenario_emissionscenario)\n")
#province = naming.split('_')[1]
province = 'AB'

housing_count = pd.read_excel('House_Counts.xlsx',header=0, nrows=1, usecols=[2,3,4], thousands=r',', 
                sheet_name= province + '_{}'.format(naming.split('_')[1]))
applliance_load = pd.read_csv('avgALP_armstrong2.csv', index_col=0, header=0) #Watt-hours

os.chdir(pwd)
os.chdir('demand_projection') #this can be changed or asked
demand_file_name = 'DESSTINEE Electricity Profiles for {}.xlsb'.format(naming.split('_')[1])
demand_proj = pd.read_excel(demand_file_name, sheet_name=province,
    header=[0], usecols=['Commercial', 'Industrial', 'Road', 'Rail'], skiprows=[0])#, skipfooter=1)

os.chdir(pwd)
directories = glob('archetypes_base/*/', recursive=True)


sen = os.path.join(pwd, naming)
os.makedirs(sen, exist_ok=True)

loads = pd.DataFrame()

for dir in directories:
    #os.chdir(f"/mnt/c/users/smoha/documents/{dir[:-1]}_{naming}/reports")
    os.chdir(f'{dir}reports')
    shutil.copy('export_meterto_csv_report_CoolingElectricity_Hourly.csv', sen + '/' + dir.split('/')[1] + '_cooling_hourly.csv')
    temp = pd.read_csv('export_meterto_csv_report_CoolingElectricity_Hourly.csv', index_col=0, parse_dates=True, header=0)
    loads[dir.split('/')[1]+'_cooling'] = temp.iloc[:,0]
    shutil.copy('export_meterto_csv_report_HeatingElectricity_Hourly.csv', sen + '/' + dir.split('/')[1] + '_heating_hourly.csv')
    temp = pd.read_csv('export_meterto_csv_report_HeatingElectricity_Hourly.csv', index_col=0, parse_dates=True, header=0)
    loads[dir.split('/')[1]+'_heating'] = temp.iloc[:,0]

    loads[dir.split('/')[1]+'_total'] = loads[dir.split('/')[1]+'_cooling'] + loads[dir.split('/')[1]+'_heating']

    os.chdir(pwd)


applliance_load.index = pd.to_datetime(loads.index)

loads['total_appliance'] = ((1/(1000000))*sum(housing_count.loc[0,:]))*applliance_load['load']
loads['residential_heating'] = (1/(3600*1000000))*(sum(housing_count[dir.split('/')[1]][0] 
    *loads[dir.split('/')[1]+'_heating'] for dir in directories))
loads['residential_cooling'] = (1/(3600*1000000))*(sum(housing_count[dir.split('/')[1]][0] 
    *loads[dir.split('/')[1]+'_cooling'] for dir in directories))
loads['residential_total'] = (1/(3600*1000000))*(sum(housing_count[dir.split('/')[1]][0] 
    *loads[dir.split('/')[1]+'_total'] for dir in directories)) \
    + loads['total_appliance']

demand_proj.index = pd.to_datetime(loads.index)
loads['commercial'] = 1000*demand_proj['Commercial']
loads['industrial'] = 1000*demand_proj['Industrial']
loads['road'] = 1000*demand_proj['Road']
loads['rail'] = 1000*demand_proj['Rail']

loads['demand'] = loads['residential_total'] + loads['commercial'] + loads['industrial'] + loads['road'] + loads['rail']

os.chdir(cwd)

copied = pwd + f'/automation/{province}_Demand_Real_Forecasted_mod.xlsx'
shutil.copy(pwd + f'/silver_prep/{province}_Demand_Real_Forecasted.xlsx', copied)

demand_base = pd.read_excel(copied, sheet_name='Province_Total_Real', index_col=0)
loads.index = pd.to_datetime(demand_base.index, dayfirst= True, unit='h')

now = time.strftime("%d-%m-%Y %H-%M-%S")

loads_shortcut = pd.DataFrame()
loads_shortcut['base_load'] = demand_base['demand']*1.05 - loads['residential_total']
loads_shortcut['residential'] = loads['residential_total']
loads_shortcut['demand'] = loads_shortcut['residential'] + loads_shortcut['base_load']

demand_base['demand'] = loads_shortcut['demand']
wb_target = load_workbook(copied, data_only=False)
del wb_target['Province_Total_Real']
writer = pd.ExcelWriter(copied, engine='openpyxl')
writer.book = wb_target
demand_base.to_excel(writer, sheet_name='Province_Total_Real', index=True)
writer.save()
wb_target.close()
shutil.copy(copied, sen + '/' + f'{province}_Demand_Real_Forecasted_shortcut_' + sen.split('/')[-2] + '.xlsx')

demand_base['demand'] = loads['demand']
wb_target = load_workbook(copied, data_only=False)
del wb_target['Province_Total_Real']
writer = pd.ExcelWriter(copied, engine='openpyxl')
writer.book = wb_target
demand_base.to_excel(writer, sheet_name='Province_Total_Real', index=True)
writer.save()
wb_target.close()
shutil.copy(copied, sen + '/' + f'{province}_2050_{naming.split("_")[0]}_{naming.split("_")[1]}_{naming.split("_")[2]}_Demand_Real_Forecasted' + '.xlsx')

totals = pd.DataFrame(loads[['residential_total', 'commercial', 'industrial', 'road', 'demand']].sum())
totals = totals.transpose()
loads = pd.concat([loads, totals], ignore_index=True)
loads.index = pd.to_datetime(loads.index, dayfirst= True, unit='h')
'''index_list = loads.index.to_list()
totals = index_list.index(8761)
index_list[totals] = 'totals'
loads.index = index_list'''
#loads.to_csv(f'loads_{now}.csv')
loads.to_csv(sen + '/' + f'loads_' + sen.split('/')[-1] + '.csv')
loads_shortcut.to_csv(sen + '/' + f'loads_shourtcut_' + sen.split('/')[-1] + '.csv')