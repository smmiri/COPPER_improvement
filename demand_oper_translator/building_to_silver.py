from glob import glob
from json import load
from openpyxl import Workbook, load_workbook
import os
import pandas as pd
import pyperclip
import shutil
import time

cwd = os.getcwd()
pwd = os.path.dirname(cwd)

naming = input("Input the iteration number province, and scenario in the correct format:\n(iteration_province_scenario in the format, iteration_elecscenario_emissionscenario)\n")
#province = naming.split('_')[1]
province = 'AB'
pyperclip.copy(naming)

iter = naming.split('_')[0]
for m in iter:
    if m.isdigit():
        iter_num = m
m = int(m)
scen = naming.split('_')[1] + '_' + naming.split('_')[2]

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
    os.chdir(f"/mnt/c/users/smoha/documents/archetypes_base/report")
    #os.chdir(f'{dir}reports')
    shutil.copy2(f'{dir.split("/")[1]}_{naming}_CoolingElectricity.csv' ,sen)
    temp = pd.read_csv(f'{dir.split("/")[1]}_{naming}_CoolingElectricity.csv', index_col=0, parse_dates=True, header=0)
    loads[dir.split('/')[1]+'_cooling'] = temp.iloc[:,0]
    shutil.copy2(f'{dir.split("/")[1]}_{naming}_HeatingElectricity.csv' ,sen)
    temp = pd.read_csv(f'{dir.split("/")[1]}_{naming}_HeatingElectricity.csv', index_col=0, parse_dates=True, header=0)
    loads[dir.split('/')[1]+'_heating'] = temp.iloc[:,0]

    loads[dir.split('/')[1]+'_total'] = loads[dir.split('/')[1]+'_cooling'] + loads[dir.split('/')[1]+'_heating']

    os.chdir(pwd)

particip = 0.25

applliance_load.index = pd.to_datetime(loads.index)

loads['total_appliance'] = ((1/(1000000))*sum(housing_count.loc[0,:]))*applliance_load['load']
loads['residential_heating'] = (1/(3600*1000000))*(sum(housing_count[dir.split('/')[1]][0] 
    *loads[dir.split('/')[1]+'_heating'] for dir in directories))
loads['residential_cooling'] = (1/(3600*1000000))*(sum(housing_count[dir.split('/')[1]][0] 
    *loads[dir.split('/')[1]+'_cooling'] for dir in directories))

if iter == 'iter0':
    loads['residential_total'] = loads['residential_heating'] + loads['residential_cooling']
    demand_prev = f'{pwd}/automation/AB_Demand_Real_Forecasted.xlsx'
else:
    if len(naming.split('_')) > 3:
        demand_prev = f"{pwd}/iter{m-1}_{scen}_{naming.split('_')[3]}/AB_2050_iter{m-1}_{scen}_Demand_Real_Forecasted.xlsx"
        loads_prev = pd.read_csv(f"{pwd}/iter{m-1}_{scen}_{naming.split('_')[3]}/loads_iter{m-1}_{scen}.csv", index_col=0)
    else:
        demand_prev = f"{pwd}/iter{m-1}_{scen}/AB_2050_iter{m-1}_{scen}_Demand_Real_Forecasted.xlsx"
        loads_prev = pd.read_csv(f"{pwd}/iter{m-1}_{scen}/loads_iter{m-1}_{scen}.csv", index_col=0)
    loads_prev.drop(loads_prev.tail(1).index, inplace=True)
    loads_prev.index = loads.index

    loads['residential_total'] = (loads['residential_cooling'] + loads['residential_heating'])*(particip) +\
        (loads_prev['residential_total'])*(1-particip)


demand_proj.index = pd.to_datetime(loads.index)
loads['commercial'] = 1000*demand_proj['Commercial']
loads['industrial'] = 1000*demand_proj['Industrial']
loads['road'] = 1000*demand_proj['Road']
loads['rail'] = 1000*demand_proj['Rail']

loads['demand'] = loads['residential_total'] + loads['commercial'] + loads['industrial'] + loads['road'] + loads['rail'] + loads['total_appliance']

os.chdir(cwd)

demand_new  = f'{sen}/AB_2050_{iter}_{scen}_Demand_Real_Forecasted.xlsx'
demand_base = pd.read_excel(demand_prev, sheet_name='Province_Total_Real', index_col=0)
loads.index = pd.to_datetime(demand_base.index, dayfirst= True, unit='h')
shutil.copy(demand_prev, demand_new)


now = time.strftime("%d-%m-%Y %H-%M-%S")

#changed_sp_prev = pd.read_csv(f'{pwd}/iter{m-1}_{scen}/changed_sp_iter{m-1}_{scen}.csv', index_col=0)
#changed_sp_prev.index = demand_base.index

#demand_base['demand'].loc[changed_sp_prev.loc[changed_sp_prev['changed']==True].index] = loads['demand']
demand_base['demand'] = loads['demand']
wb_target = load_workbook(demand_new, data_only=False)
del wb_target['Province_Total_Real']
writer = pd.ExcelWriter(demand_new, engine='openpyxl')
writer.book = wb_target
demand_base.to_excel(writer, sheet_name='Province_Total_Real', index=True)
writer.save()
wb_target.close()

totals = pd.DataFrame(loads[['residential_total', 'commercial', 'industrial', 'road', 'demand']].sum())
totals = totals.transpose()
loads = pd.concat([loads, totals], ignore_index=True)
loads.index = pd.to_datetime(loads.index, dayfirst= True, unit='h')
'''index_list = loads.index.to_list()
totals = index_list.index(8761)
index_list[totals] = 'totals'
loads.index = index_list'''
#loads.to_csv(f'loads_{now}.csv')
loads.to_csv(f"{sen}/loads_{naming.split('_')[0]}_{naming.split('_')[1]}_{naming.split('_')[2]}.csv")