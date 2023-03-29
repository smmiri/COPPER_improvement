import glob
import numpy as np
import os
import pandas as pd
import shutil

# Function to fix the path to the SILVER output folder in case of Windows
def fixpath(path):
    if path.startswith("C:"): return "/mnt/c/" + path.replace("\\", "/")[3:]
    else:
        pass
    return path

#silver_output_dir = fixpath(input('Insert the path to the folder containing SILVER results:\n'))
scen = input('Scenario:\n')
ext = input('Extension for the scenario:\n')
silver_output_dir = fixpath(r"C:\SILVER_BC_Cascade\SILVER_Data\Model Results\BC_Cascade_{}".format(scen))
scen_dir = fixpath(r"C:\Users\smoha\OneDrive - University of Victoria\Project\tasks\linkage_hydro\scen_outputs\BC_cascade_{}_{}".format(scen, ext))
os.makedirs(scen_dir, exist_ok=True)
files = os.listdir(silver_output_dir)
for fname in files:
    shutil.copy2(os.path.join(silver_output_dir, fname), scen_dir)
os.chdir(scen_dir)


lmp = False
#lmp = True if input('Are you analysing lmp? (y/n)\t') == 'y' else False

extension = 'csv'
vre_filenames = [i for i in glob.glob('Available_VRE_generation*.{}'.format(extension))]
combined_available_vre = pd.concat([pd.read_csv(f, index_col=0) for f in vre_filenames], axis=0)
combined_available_vre.index = pd.to_datetime(combined_available_vre.index, unit='h' ,dayfirst=True, origin='2050-01-01')
combined_available_vre = combined_available_vre.drop(columns='date')

# Function to check if a row is a date to drop non-date rows form UC output
def check_row(a):
    try:
        np.datetime64(a)
    except:
        return False
    else:
        return True
        
uc_filenames = [i for i in glob.glob('UC_Results_*.{}'.format(extension))]
header = list(pd.read_csv(uc_filenames[0], header=None, index_col=0).loc['name',:])
header = [x for x in header if str(x) != 'nan' and x != 'name']
gen_types = list(dict.fromkeys([x.split('_')[0] for x in header]))
header.insert(0,'date')
combined_uc = pd.concat([pd.read_csv(f, index_col=0, engine='python') for f in uc_filenames])
combined_uc.dropna(axis=0, subset=[combined_uc.columns[1]], inplace=True)
combined_uc = combined_uc.loc[combined_uc.index.map(check_row),:]
combined_uc = combined_uc.apply(pd.to_numeric)
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
max_col = lambda x: abs(max(x.max(), x.min(), key=abs))
combined_lf.loc['max'] = combined_lf.apply(max_col, axis=0)

congestion = pd.DataFrame(index=combined_lf.index, columns=range(len(combined_lf.columns)))
congestion.columns = pd.MultiIndex.from_frame(header)
pmax = combined_lf.loc['max']

for column in combined_lf.columns:
    congestion[column] = np.where(combined_lf[column] > 0.95*pmax[column], 1, 0) # 0.95 is the threshold for congestion, could be changed
congestion['Any congested lines'] = congestion.any(axis=1)
congestion.loc['total'] = congestion.sum()
congestion.drop('max', inplace=True)

curtailment_rate = pd.DataFrame(index=combined_uc_vre.index)
curtailment_rate['Total Available Wind'] = combined_available_vre[list(combined_uc.filter(regex='^Wind'))].sum(axis=1)
curtailment_rate['Total Dispatched Wind'] = combined_uc_vre[list(combined_uc.filter(regex='^Wind'))].sum(axis=1)
curtailment_rate['Total Curtailed Wind'] = (curtailment_rate['Total Available Wind']-curtailment_rate['Total Dispatched Wind'])
curtailment_rate['Total Available Solar'] = combined_available_vre[list(combined_uc.filter(regex='^Solar'))].sum(axis=1)
curtailment_rate['Total Dispatched Solar'] = combined_uc_vre[list(combined_uc.filter(regex='^Solar'))].sum(axis=1)
curtailment_rate['Total Curtailed Solar'] = (curtailment_rate['Total Available Solar']-curtailment_rate['Total Dispatched Solar'])
curtailment_rate['Total Available VRE'] = combined_available_vre[list(combined_uc.filter(regex='^Wind|^Solar'))].sum(axis=1)
curtailment_rate['Total Dispatched VRE'] = combined_uc_vre[list(combined_uc.filter(regex='^Wind|^Solar'))].sum(axis=1)
curtailment_rate['Total Curtailed VRE'] = (curtailment_rate['Total Available VRE']-curtailment_rate['Total Dispatched VRE'])
curtailment_rate.loc['curtailed percentage', 'Total Curtailed Wind'] = round(100*curtailment_rate['Total Curtailed Wind'].sum()/curtailment_rate['Total Available Wind'].sum(),3)
curtailment_rate.loc['curtailed percentage', 'Total Curtailed Solar'] = round(100*curtailment_rate['Total Curtailed Solar'].sum()/curtailment_rate['Total Available Solar'].sum(),3)
curtailment_rate.loc['curtailed percentage', 'Total Curtailed VRE'] = round(100*curtailment_rate['Total Curtailed VRE'].sum()/curtailment_rate['Total Available VRE'].sum(),3)



curtailment = combined_available_vre - combined_uc_vre
curtailment.loc['total'] = curtailment.sum()
curtailment.loc['curtailed percentage'] = round(100*curtailment.loc['total']/combined_available_vre.sum(),3)

curtailment = pd.concat([curtailment, curtailment_rate], axis=1)

curtailment_rate.loc['total'] = curtailment_rate.sum()
combined_available_vre.loc['total'] = combined_available_vre.sum()
combined_uc_vre.loc['total'] = combined_uc_vre.sum()
combined_uc.loc['total'] = combined_uc.sum()

total_dispatch = pd.DataFrame()
for gen_type in gen_types:
    total_dispatch[gen_type] = combined_uc[list(combined_uc.filter(regex=f'^{gen_type}'))].sum(axis=1)
if 'cascade' in total_dispatch.columns:  
    total_dispatch['hydro'] = total_dispatch['hydro'] + total_dispatch['cascade']
    total_dispatch = total_dispatch.drop(columns=['cascade'])

if lmp:
    #LMP analysis
    all_filenames = [i for i in glob.glob('LMP*.{}'.format(extension))]
    combined_lmp = pd.concat([pd.read_csv(f, index_col=0) for f in all_filenames], axis=1)

    lmp_hourly = combined_lmp.T
    lmp_hourly.reset_index(drop=True, inplace=True)
    lmp_hourly.index = pd.to_datetime(lmp_hourly.index, unit='h', dayfirst=True, origin='2050-01-01')
    lmp_hourly['mean'] = lmp_hourly.mean(axis=1)

analysis = pd.DataFrame(index=curtailment_rate.index)
analysis['Curtailed VRE'] = curtailment_rate['Total Curtailed VRE']
if lmp:
    analysis['LMP(mean)'] = lmp_hourly['mean']
analysis['Load'] = total_dispatch.sum(axis=1)

#output_dir = fixpath(r"C:\Users\smoha\OneDrive - University of Victoria\Project\tasks\linkage_hydro\pre_analysis") #change this to your output directory
#os.chdir(output_dir)

writer = pd.ExcelWriter(f'analysis_{silver_output_dir.split("/")[-1]}.xlsx', engine='xlsxwriter')
analysis.to_excel(writer, sheet_name="Analysis")
combined_uc.to_excel(writer, sheet_name="UC Results")
combined_uc_vre.to_excel(writer, sheet_name="UC VRE Results")
combined_lf.to_excel(writer, sheet_name="Line Flow")
combined_available_vre.to_excel(writer, sheet_name="Available VRE")
total_dispatch.to_excel(writer, sheet_name="Total Dispatch")
if lmp:
    lmp_hourly.to_excel(writer, sheet_name="LMP Results")
curtailment.to_excel(writer, sheet_name="Curtailment Details")
congestion.to_excel(writer, sheet_name="Congestion Analysis")
writer.close()