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
'''to add calcs for congestion'''

total_avail = combined_available_vre.sum()
total_avail.name = 'total'
combined_available_vre.loc[len(combined_available_vre),] = total_avail
total_gen = combined_uc_vre.sum()
total_gen.name = 'total'
combined_uc_vre.loc[len(combined_uc_vre)] = total_gen
curtailment = 100* (1 - combined_uc_vre / combined_available_vre)
curtailment.loc[len(curtailment)] = total_avail
curtailment.loc[len(curtailment)] = total_gen
curtailment.index.values[-3:] = ['curtailed percentage', 'available generation', 'actual generation']

combined_uc.loc[len(combined_uc)] = combined_uc.sum()
combined_uc.index.values[-1] = 'sum'
combined_uc_vre[len(combined_uc_vre)] = combined_uc_vre.sum()
combined_uc_vre.index.values[-1] = 'sum'

combined_uc.to_csv('combined_uc.csv')
combined_uc_vre.to_csv("combined_uc_vre.csv", encoding='utf-8-sig')
curtailment.to_excel('curtailment.xlsx')
combined_lf.to_csv("combined_lineflow.csv", encoding='utf-8-sig')
combined_available_vre.to_csv('combined_available_vre.csv')