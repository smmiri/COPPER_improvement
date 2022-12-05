import glob
import pandas as pd
import os

cwd = os.getcwd()
pwd = os.path.dirname(cwd)
dir = input('Insert the folder containing SILVER results:\n')
os.chdir(f'{pwd}/{dir}')

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

try:
    shutil.copy2(f'{pwd}/iter{m-1}_{scen}/pmax.xlsx', f'{pwd}/{dir}')
except:
    print('This is the first iteration')

try:
    pmax = pd.read_excel('pmax.xlsx', index_col=0, header=None)
except:
    print('Provide PMAX data for this scenario')

pmax = pmax.drop(['from','to'])
pmax.columns = pd.MultiIndex.from_frame(header)
pmax = pd.concat([pmax]*1440)
pmax.index = combined_lf.index

for column in combined_lf.columns:
    congestion[column] = np.where(combined_lf[column] > 0.95*pmax[column], 1, 0)

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

#LMP analysis
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

writer = pd.ExcelWriter(f'analysis_{dir}.xlsx', engine='xlsxwriter')
analysis.to_excel(writer, sheet_name="Analysis", encoding='UTF-8')
combined_uc.to_excel(writer, sheet_name="UC Results", encoding='UTF-8')
combined_uc_vre.to_excel(writer, sheet_name="UC VRE Results", encoding='UTF-8')
combined_lf.to_excel(writer, sheet_name="Line Flow", encoding='UTF-8')
combined_available_vre.to_excel(writer, sheet_name="Available VRE", encoding='UTF-8')
lmp_hourly.to_excel(writer, sheet_name="LMP Results", encoding='UTF-8')
curtailment.to_excel(writer, sheet_name="Curtailment Details", encoding='UTF-8')
congestion.to_excel(writer, sheet_name="Congestion Analysis", encoding='UTF-8')
writer.save()
