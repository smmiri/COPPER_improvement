import glob
import shutil
import numpy as np
from openpyxl import load_workbook
import os
import pandas as pd

cwd = os.getcwd()
pwd = os.path.dirname(cwd)
dir = input("Input the iteration number province, and scenario in the correct format:\n(iteration_province_scenario in the format, iteration_elecscenario_emissionscenario)\n")
iter = dir.split('_')[0]
for m in iter:
    if m.isdigit():
        iter_num = m
m = int(m)
scen = dir.split('_')[1] + '_' + dir.split('_')[2]
src = f"/mnt/c/silver_lmp/silver_data/model results/ab_2050_iter{m-1}_{scen}"
files = os.listdir(src)
for fname in files:
    shutil.copy2(os.path.join(src, fname), f'{pwd}/iter{m-1}_{scen}')

shutil.copy2(f'{pwd}/pmax.xlsx', f'{pwd}/iter{m-1}_{scen}')

os.chdir(pwd)
os.makedirs(os.path.join(pwd, dir), exist_ok=True)
os.chdir(f'{pwd}/iter{m-1}_{scen}')


###Post processing SILVER results###
extension = 'csv'
vre_filenames = [i for i in glob.glob('Available_VRE_generation*.{}'.format(extension))]
combined_available_vre = pd.concat([pd.read_csv(f, index_col=0) for f in vre_filenames], axis=0)
combined_available_vre.index = pd.to_datetime(combined_available_vre.index, unit='h' ,dayfirst=True, origin='2050-01-01')
combined_available_vre = combined_available_vre.drop(columns='date')

uc_filenames = [i for i in glob.glob('UC_Results_*.{}'.format(extension))]
header = list(pd.read_csv(uc_filenames[0], header=None).loc[14,:])
header = [x for x in header if str(x) != 'nan' and x != 'name']
header.insert(0,'date')
combined_uc = pd.concat([pd.read_csv(f, skiprows=[i for i in range(1,30)], index_col=0,skipfooter=749, engine='python') for f in uc_filenames])
combined_uc.reset_index(inplace=True)
combined_uc.index = pd.to_datetime(combined_uc.index, unit='h', dayfirst=True, origin='2050-01-01')
combined_uc = combined_uc.drop(columns=['Total','dr'])
combined_uc.columns = header
combined_uc = combined_uc.drop(columns=['date'])
combined_uc_vre = combined_uc[list(combined_uc.filter(regex='^Wind|^Solar'))]

lf_filenames = [i for i in glob.glob('Line_Flow_*.csv')]
header = pd.DataFrame(pd.read_csv(lf_filenames[0]))
header = header.iloc[:,[0,1]]
combined_lf = pd.concat([(pd.read_csv(f, header=None)) for f in lf_filenames], axis=1)
combined_lf = combined_lf.transpose()
combined_lf.index = combined_lf[0]
combined_lf = combined_lf.drop(columns=[0])
combined_lf = combined_lf.drop(index=['from','to'])
combined_lf.reset_index(inplace=True)
combined_lf = combined_lf.drop(columns=[0])
combined_lf.index = pd.to_datetime(combined_lf.index, unit='h', dayfirst=True, origin='2050-01-01')
combined_lf.columns = pd.MultiIndex.from_frame(header)
combined_lf = combined_lf.astype(float)

congestion = pd.DataFrame(index=combined_lf.index, columns=range(83))
congestion.columns = pd.MultiIndex.from_frame(header)
pmax = pd.read_excel('pmax.xlsx', index_col=0, header=None)
pmax = pmax.drop(['from','to'])
pmax.columns = pd.MultiIndex.from_frame(header)
pmax = pd.concat([pmax]*1440)
pmax.index = combined_lf.index

for column in combined_lf.columns:
    congestion[column] = np.where(combined_lf[column] > 0.85*pmax[column], 1, 0)

congestion['Any congested lines'] = congestion.any(axis=1)

total_avail = combined_available_vre.sum()
total_avail.name = 'total'
combined_available_vre.loc[len(combined_available_vre),] = total_avail
total_gen = combined_uc_vre.sum()
total_gen.name = 'total'
combined_uc_vre.loc[len(combined_uc_vre)] = total_gen

curtailment_rate = pd.DataFrame(index=combined_uc_vre.index)
curtailment_rate['Total Available Wind'] = combined_available_vre[list(combined_uc.filter(regex='^Wind'))].sum(axis=1)
curtailment_rate['Total Dispatched Wind'] = combined_uc_vre[list(combined_uc.filter(regex='^Wind'))].sum(axis=1)
curtailment_rate.loc[len(curtailment_rate)] = curtailment_rate.sum()
curtailment_rate.index.values[-1] = 'Total'
curtailment_rate['Total Curtailed Wind'] = 100*(1-curtailment_rate['Total Dispatched Wind']/curtailment_rate['Total Available Wind'])
curtailment_rate = curtailment_rate.drop(1440)

curtailment = 100* (1 - combined_uc_vre / combined_available_vre)
curtailment.loc[len(curtailment)] = total_avail
curtailment.loc[len(curtailment)] = total_gen
curtailment = pd.concat([curtailment, curtailment_rate], axis=1)
curtailment.index.values[-3:] = ['curtailed percentage', 'available generation', 'actual generation']


combined_uc.loc[len(combined_uc)] = combined_uc.sum()
combined_uc.index.values[-1] = 'sum'
combined_uc_vre[len(combined_uc_vre)] = combined_uc_vre.sum()
combined_uc_vre.index.values[-1] = 'sum'

###Translating LMP results into setpoint measures###
all_filenames = [i for i in glob.glob('LMP*.{}'.format(extension))]
combined_lmp = pd.concat([pd.read_csv(f, index_col=0) for f in all_filenames], axis=1)

lmp_hourly = combined_lmp.T
lmp_hourly.reset_index(drop=True, inplace=True)
lmp_hourly.index = pd.to_datetime(lmp_hourly.index, unit='h', dayfirst=True, origin='2050-01-01')
lmp_hourly['mean'] = lmp_hourly.mean(axis=1)

analysis = pd.DataFrame(index=curtailment_rate.index)
analysis['Curtailed Wind'] = curtailment_rate['Total Curtailed Wind']
analysis['LMP(mean)'] = lmp_hourly['mean']
analysis['Load'] = combined_uc.sum(axis=1)

writer = pd.ExcelWriter(f'{pwd}/iter{m-1}_{scen}/analysis_iter{m-1}_{scen}.xlsx', engine='xlsxwriter')
analysis.to_excel(writer, sheet_name="Analysis", encoding='UTF-8')
combined_uc.to_excel(writer, sheet_name="UC Results", encoding='UTF-8')
combined_uc_vre.to_excel(writer, sheet_name="UC VRE Results", encoding='UTF-8')
combined_lf.to_excel(writer, sheet_name="Line Flow", encoding='UTF-8')
combined_available_vre.to_excel(writer, sheet_name="Available VRE", encoding='UTF-8')
combined_lmp.to_excel(writer, sheet_name="LMP Results", encoding='UTF-8')
curtailment.to_excel(writer, sheet_name="Curtailment Details", encoding='UTF-8')
congestion.to_excel(writer, sheet_name="Congestion Analysis", encoding='UTF-8')
writer.save()

load = pd.read_csv(f'{pwd}/iter{m-1}_{scen}/loads_iter{m-1}_{scen}.csv', index_col=0)

load = load[['residential_cooling', 'residential_heating']].reset_index(drop=True)
load.index = pd.to_datetime(load.index, unit='h', dayfirst=True, origin='2050-01-01')

lmp_daily = pd.DataFrame()
el_upper = 0.75
el_lower = 0.25

lmp_daily['sp_max'] = lmp_hourly['mean'].groupby(np.arange(len(lmp_hourly))//24).max()
lmp_daily['sp_min'] = lmp_hourly['mean'].groupby(np.arange(len(lmp_hourly))//24).min()
lmp_daily['hpt'] = lmp_daily['sp_min'] + el_upper*(lmp_daily['sp_max']-lmp_daily['sp_min'])
lmp_daily['lpt'] = lmp_daily['sp_min'] + el_lower*(lmp_daily['sp_max']-lmp_daily['sp_min'])
lmp_daily.reset_index(drop=True, inplace=True)
lmp_daily.index = pd.to_datetime(lmp_daily.index, unit='D', dayfirst=True, origin='2050-01-01')
#lmp_daily.to_csv('temp.csv')


#testing with only 5 months
#lmp_hourly = lmp_hourly.loc[lmp_hourly.index.month == [1,2,3,4]]

Rubystringhtg = 'ems_htg_setpoint_prg.addLine'
Rubystringclg = 'ems_clg_setpoint_prg.addLine'

max_sp_winter = 24
min_sp_winter = 20
mean_sp_winter = 21
max_sp_summer = 27
min_sp_summer = 23
mean_sp_summer = 25

lastmonth = 0
lastday = 0
setpoint_counter = 0
monthly_hours = 0

red_imp = 0.96
inc_imp = 1.04

changed_sp = pd.DataFrame(index = combined_lf.index)
changed_sp['changed'] = np.nan

#changed_sp_prev = pd.read_csv(f'{pwd}/iter{m-1}_{scen}/changed_sp_iter{m-1}_{scen}.csv')
loads_prev = pd.read_csv(f'{pwd}/iter{m-1}_{scen}/loads_iter{m-1}_{scen}.csv', index_col=0)
#changed_sp['changed'] = changed_sp_prev['changed'].to_numpy()

loads_prev.drop(loads_prev.tail(1).index, inplace=True)
loads_prev.reset_index(drop=True, inplace=True)
loads_prev.index = pd.to_datetime(loads_prev.index, unit='h', dayfirst=True, origin='2050-01-01')
loads = loads_prev

for index,values in lmp_hourly.iterrows():
    
    currentmonth = index.month
    currentday = index.day
    hour = index.hour

    #separate seasons, winter, based on the load calculation files out of base load
    if load.loc[(load.index.hour == hour) & (load.index.day == currentday) & (load.index.month == currentmonth), 'residential_heating'][0] != 0:
        #and changed_sp['changed'].loc[index] != True:

        if  (values['mean'] > lmp_daily.loc[(lmp_daily.index.day == currentday) & (lmp_daily.index.month == currentmonth), 'hpt'][0] or  
             values['mean'] < lmp_daily.loc[(lmp_daily.index.day == currentday) & (lmp_daily.index.month == currentmonth), 'lpt'][0]):
             
             changed_sp.loc[index, 'changed'] = True
             
                     
        #decrease the load, if higher than HPT
        if values['mean'] > lmp_daily.loc[(lmp_daily.index.day == currentday) & (lmp_daily.index.month == currentmonth), 'hpt'][0]:# and \
            #congestion.loc[index, congestion.columns.get_level_values(0)=='Any congested lines'].values[0] and \
            #curtailment_rate['Total Curtailed Wind'].loc[index] < 5:# and changed_sp['changed'].loc[index] != True:

            loads.loc[index, 'residential_heating'] = loads_prev.loc[index, 'residential_heating'] * red_imp


        #increase the load, if lower than LPT
        elif values['mean'] < lmp_daily.loc[(lmp_daily.index.day == currentday) & (lmp_daily.index.month == currentmonth), 'lpt'][0]:# and \
            #congestion.loc[index, congestion.columns.get_level_values(0)=='Any congested lines'].values[0]==False and \
            #curtailment_rate['Total Curtailed Wind'].loc[index] >= 5:# and changed_sp['changed'].loc[index] != True:

            loads.loc[index, 'residential_heating'] = loads_prev.loc[index, 'residential_heating'] * inc_imp

    lastmonth = currentmonth
    lastday = currentday

changed_sp.to_csv(f'{pwd}/{dir}/changed_sp_{dir}.csv')

loads['residential_total'] = loads['residential_heating'] + loads['residential_cooling'] + loads['total_appliance']
loads['demand'] = loads['residential_total'] + loads['commercial'] + loads['industrial'] + loads['road'] + loads['rail']

copied = f'{pwd}/iter{m}_{scen}/AB_2050_iter{m}_{scen}_Demand_Real_Forecasted.xlsx'
shutil.copy(f'{pwd}/iter{m-1}_{scen}/AB_2050_iter{m-1}_{scen}_Demand_Real_Forecasted.xlsx', copied)

demand_base = pd.read_excel(copied, sheet_name='Province_Total_Real', index_col=0)
loads.index = pd.to_datetime(demand_base.index, dayfirst= True, unit='h')

demand_base['demand'] = loads['demand']
wb_target = load_workbook(copied, data_only=False)
del wb_target['Province_Total_Real']
writer = pd.ExcelWriter(copied, engine='openpyxl')
writer.book = wb_target
demand_base.to_excel(writer, sheet_name='Province_Total_Real', index=True)
writer.save()
wb_target.close()

totals = pd.DataFrame(loads[['residential_total', 'commercial', 'industrial', 'road', 'demand']].sum())
totals = totals.transpose()
loads = pd.concat([loads, totals], ignore_index=True)
loads.reset_index(inplace=True, drop=True)
loads.index = pd.to_datetime(loads.index, dayfirst= True, unit='h', origin='2050-01-01')
loads.to_csv(f'{pwd}/iter{m}_{scen}/loads_iter{m}_{scen}.csv')