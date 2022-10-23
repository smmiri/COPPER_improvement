from glob import glob
from json import load
import os
import matplotlib.pyplot as plt
from numpy import NaN
import pandas as pd
import seaborn as sns

sns.set_theme()


cwd = os.getcwd()
pwd = os.path.dirname(cwd)

dir = 'scen_output'
os.chdir(pwd)
iters = glob(f'{dir}/*/', recursive=True)

ap_scen = []
for iter in iters:
    scen = iter.split('/')[1]
    ap_scen.append(scen)
scens = pd.DataFrame(columns = ['scen', 'elec_scen', 'emit_scen', 'load_chg', 'dr_method'])
scens['scen'] = ap_scen


scens[['elec_scen', 'emit_scen', 'load_chg', 'dr_method']] = scens['scen'].str.split('_', expand=True)
scens['load_chg'] = list(scens.iloc[:, 3].astype(str).str.extractall('(\d+)').fillna('').sum(axis=1).astype(int))


os.chdir(dir)

print(scens['scen'])
scen_no = input("please insert scenario number or multiple for compare separated by ',' or say all:\n")
if scen_no == 'all':
    scen_num = scens.index.values.tolist()
else:
    scen_num = str(scen_no).split(',')
    scen_num = list(map(int, scen_num))


iters_no = []
curts_iter = pd.DataFrame()
e_sav_iters = pd.DataFrame()
costs_iter = pd.DataFrame()
emissions_iter = pd.DataFrame()
iters_list = ['iter0', 'iter1', 'iter2', 'iter3', 'iter4', 'iter5', 'iter6', 'iter7', 'iter8', 'iter9']
costs_iter['iters'] = iters_list
curts_iter['iters'] = iters_list
e_sav_iters['iters'] = iters_list
emissions_iter['iters'] = iters_list
gen_types = {'Wind':'Wind', 'Solar':'Solar','Biomass':'Biomass', 'NG':'Gas', 'hydro':'Hydro'}
gen_price = {'Wind':0.00001, 'Solar':.00001,'Biomass':5.06, 'NG':2.67, 'hydro':1.46}
gen_titles = ['Wind', 'Solar','Biomass', 'Gas', 'Hydro']
dispatch_iters = pd.DataFrame(columns=iters_list, index=gen_types.values())



for index, row in scens.loc[scen_num].iterrows():
    load_curve = pd.DataFrame()
    curts = []
    e_savs = []
    total_costs = []
    emissions = []
    scen=row['scen']
    os.chdir(scen)
    iters = glob('*/', recursive=True)
    iters_no = pd.DataFrame(iters)
    iters_no[['iters_dir','drop']] = iters_no.iloc[:,0].str.split('/',expand=True)
    iters_no[['iters', 'elec_scen', 'emit_scen']] = iters_no.iloc[:,0].str.split('_',expand=True)
    for iter in iters_no['iters_dir']:
        os.chdir(iter)
        iter_name = iter.split('_')[0]
        general = pd.read_excel(f'analysis_{iter}.xlsx', sheet_name='Analysis', index_col=0)
        curt = general.loc['Total', 'Curtailed Wind'].item()
        curts.append(curt)
        
        load_curve[iter_name] = general['Load']
        
        e_sav = general['Load'].sum()
        e_savs.append(e_sav)
        
        uc_results = pd.read_excel(f'analysis_{iter}.xlsx', sheet_name='UC Results', index_col=0)
        uc_results = uc_results.loc['sum']
        total_cost = round((uc_results[['Biomass' in s for s in uc_results.index]].sum()*5.06 +\
                      uc_results[['Wind' in s for s in uc_results.index]].sum()*.01 +\
                      uc_results[['Solar' in s for s in uc_results.index]].sum()*.01 +\
                      uc_results[['hydro' in s for s in uc_results.index]].sum()*1.46 +\
                      uc_results[['NG' in s for s in uc_results.index]].sum()*2.67)/1000 , 2)
        total_costs.append(total_cost)
        
        for gen_type in gen_types.keys():
            dispatch_iters.loc[gen_types[gen_type], iter_name] = uc_results[[gen_type in s for s in uc_results.index]].sum()

        emission = uc_results[['NG' in s for s in uc_results.index]].sum()*0.576*0.000001
        emissions.append(emission)

        os.chdir(f"{pwd}/{dir}/{scen}")


    costs_iter[scen] = pd.Series(total_costs)
    curts_iter[scen] = pd.Series(curts)
    e_sav_iters[scen] = pd.Series(e_savs)
    emissions_iter[scen] = pd.Series(emissions)
    load_curve.drop(load_curve.tail(1).index, inplace=True)
    load_curve.index = general.index[0:1440]
    
    plot_title = f"Curtailed Wind in % - Scenario: {row['elec_scen']}, {row['emit_scen']} - DR Method: {row['dr_method']} with {row['load_chg']}% participation"
    sns.set(rc={'figure.figsize': (8,5)})
    crt_scen = sns.lineplot(data = curts_iter, x='iters', y=scen)
    crt_scen.set(xlabel= "Iterations", ylabel= "Curtailed Wind genartion (%)")
    plt.savefig(f'curtailments_{scen}.jpg', dpi=300, bbox_inches= 'tight')
    plt.close()

    e_sav_scen = sns.lineplot(data = e_sav_iters, x='iters', y=scen)
    e_sav_scen.set(xlabel= "Iterations", ylabel= "Total Energy Supplied (MWh)")
    plt.savefig(f'esavings_{scen}.jpg', dpi=300, bbox_inches= 'tight')
    plt.close()

    costs_scen = sns.lineplot(data = costs_iter, x='iters', y=scen)
    costs_scen.set(xlabel= "Iterations", ylabel= "Total Yearly Operational Costs (Thousand$)")
    plt.savefig(f'costs_{scen}.jpg', dpi=300, bbox_inches= 'tight')
    plt.close()

    emissions_scen = sns.lineplot(data = emissions_iter, x='iters', y=scen)
    emissions_scen.set(xlabel= "Iterations", ylabel= "Total emissions in million ton CO2eq")
    plt.savefig(f'emissions_{scen}.jpg', dpi=300, bbox_inches= 'tight')
    plt.close()

    dispatch_iters = dispatch_iters.transpose()
    dispatch_iters.dropna(subset='Wind', inplace=True)
    dispatch_fig = dispatch_iters.plot(kind='bar', stacked=True, color=['tab:green','yellow','grey','black','tab:blue'])
    dispatch_fig.set(xlabel='Iterations',  ylabel='Dispatched generation (MWh)')
    plt.savefig(f'dispatch_{scen}.jpg', dpi=300, bbox_inches= 'tight')
    plt.close()

    sns.set(rc={'figure.figsize': (20,5)})
    load_curve_iter = sns.lineplot(data = load_curve.iloc[24:96])
    load_curve_iter.set(xlabel= "Time", ylabel= "Load (MWh)")
    plt.savefig(f'load_curves_{scen}.jpg', dpi=300, bbox_inches= 'tight')
    plt.close()
   
    del curts
    del load_curve
    del e_savs
    del total_costs
    del emissions
    os.chdir(f'{pwd}/{dir}')

if len(scen_num) > 1:
    sns.set(rc={'figure.figsize': (8,5)})
    crt_total = sns.lineplot(data = curts_iter)
    crt_total.set(xlabel= "Iterations", ylabel= "Curtailed Wind genartion (%)")
    plt.savefig(f"curtailments_{'_'.join(str(item) for item in scen_num)}.jpg", dpi=300, bbox_inches= 'tight')
    plt.close()

    esav_total = sns.lineplot(data = e_sav_iters)
    esav_total.set(xlabel= "Iterations", ylabel= "Total Energy Supplied (MWh)")
    plt.savefig(f"esavings_{'_'.join(str(item) for item in scen_num)}.jpg", dpi=300, bbox_inches= 'tight')
    plt.close()
