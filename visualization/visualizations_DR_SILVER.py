from glob import glob
import os
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
import numpy as np
import pandas as pd
import seaborn as sns

sns.color_palette("Set2")
sns.set_style('white')
plt.rc('lines', linewidth=0.75)



cwd = os.getcwd()
pwd = os.path.dirname(cwd)

#dir = 'scen_archive'
#dir = 'scen_output'
dir = 'scen_output_badparticip'
os.chdir(pwd)
iters = glob(f'{dir}/*/', recursive=True)

ap_scen = []
for iter in iters:
    scen = iter.split('/')[1]
    ap_scen.append(scen)
scens = pd.DataFrame(columns = ['scen', 'elec_scen', 'emit_scen', 'particip', 'dr_method'])
scens['scen'] = ap_scen


scens[['elec_scen', 'emit_scen', 'particip', 'dr_method']] = scens['scen'].str.split('_', expand=True)
scens['particip'] = list(scens.iloc[:, 3].astype(str).str.extractall('(\d+)').fillna('').sum(axis=1).astype(int))


os.chdir(dir)

print(scens['scen'])
scen_no = input("Please insert scenario number or numbers separated by ',' or say 'all':\n")
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
analysis_scen = pd.DataFrame()
wind_avail = pd.DataFrame()
iters_list = ['Iteration 1', 'Iteration 2', 'Iteration 3', 'Iteration 4', 'Iteration 5', 'Iteration 6', 'Iteration 7', 'Iteration 8', 'Iteration 9']
costs_iter['iters'] = iters_list
curts_iter['iters'] = iters_list
e_sav_iters['iters'] = iters_list
emissions_iter['iters'] = iters_list
solar_out = pd.DataFrame()
solar_out['iters'] = iters_list
gen_types = {'Wind':'Wind', 'Solar':'Solar','Biomass':'Biomass', 'NG':'Gas', 'hydro':'Hydro'}
gen_price = {'Wind':0.00001, 'Solar':.00001,'Biomass':5.06, 'NG':2.67, 'hydro':1.46}
gen_titles = ['Wind', 'Solar','Biomass', 'Gas', 'Hydro']
dispatch_iters = pd.DataFrame(columns=iters_list, index=gen_types.values())


for index, row in scens.loc[scen_num].iterrows():
    iters_list = ['Iteration 1', 'Iteration 2', 'Iteration 3', 'Iteration 4', 'Iteration 5', 'Iteration 6', 'Iteration 7', 'Iteration 8', 'Iteration 9']
    analysis_table = pd.DataFrame()
    load_curve = pd.DataFrame()
    curts = []
    e_savs = []
    total_costs = []
    emissions = []
    wind_outs = []
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
        load = pd.read_csv(f'loads_{iter}.csv', index_col=0)
        load.reset_index(inplace=True)
        load.index = pd.to_datetime(load.index, unit='h', origin='2050-01-01')
        curt = general.loc['Total', 'Curtailed Wind'].item()
        curts.append(curt)

        curt_det = pd.read_excel(f'analysis_{iter}.xlsx', sheet_name='Curtailment Details', index_col=0)
        wind_out = curt_det['Total Dispatched Wind'].sum()
        analysis_table.loc[iter_name, 'Wind Output (MWh)'] = wind_out
        load_curve[iter_name] = load['demand'].iloc[0:1440]
        e_sav = load['demand'].iloc[-1]
        e_savs.append(e_sav)
        
        uc_results = pd.read_excel(f'analysis_{iter}.xlsx', sheet_name='UC Results', index_col=0)
        uc_results = uc_results.loc['sum']
        total_cost = round((uc_results[['Biomass' in s for s in uc_results.index]].sum()*5.06 +\
                      uc_results[['Wind' in s for s in uc_results.index]].sum()*.01 +\
                      uc_results[['Solar' in s for s in uc_results.index]].sum()*.01 +\
                      uc_results[['hydro' in s for s in uc_results.index]].sum()*1.46 +\
                      uc_results[['NG' in s for s in uc_results.index]].sum()*2.67)/1000 , 2)
        total_costs.append(total_cost)
        analysis_table.loc[iter_name, 'Operational Costs (M$)'] = total_cost

        emissions_analysis = round((uc_results[['NG' in s for s in uc_results.index]].sum()*0.655)/1000000 , 2)
        analysis_table.loc[iter_name, 'Emissions (MTCO2e)'] = emissions_analysis

        house_load = load.iloc[0:1439, 13]
        mp = general['LMP(mean)']
        house_cost = house_load.mul(mp, axis='index')
        analysis_table.loc[iter_name, 'Household Energy Expenses (M$)'] = house_cost.sum()/1000000

        if row['emit_scen'] == 'emit':
            vre_avail = pd.read_excel(f'analysis_{iter}.xlsx', sheet_name='Available VRE', index_col=0)
            solar_avail = vre_avail.iloc[0:1440, 0:16].sum(axis=1)
            out_vre = pd.read_excel(f'analysis_{iter}.xlsx', sheet_name='UC VRE Results', index_col=0)

        
        for gen_type in gen_types.keys():
            dispatch_iters.loc[gen_types[gen_type], iter_name] = uc_results[[gen_type in s for s in uc_results.index]].sum()

        emission = uc_results[['NG' in s for s in uc_results.index]].sum()*0.655*0.000001
        emissions.append(emission)

        os.chdir(f"{pwd}/{dir}/{scen}")

    if row['emit_scen'] == 'emit':
        solar_out[scen] = out_vre.iloc[0:1440, 0:16].sum(axis=1)
        wind_avail = curt_det['Total Available Wind']
        wind_avail.dropna(inplace=True)
        wind_avail = wind_avail + solar_avail
    else:
        wind_avail = curt_det['Total Available Wind']
        wind_avail.dropna(inplace=True)
    
    total_costs[-1] = total_costs[-2] #fixing last value for the reduced load cases
    curts[-1] = curts[-2] #fixing last value for the reduced load cases
    load_curve.iloc[:,-1] = load_curve.iloc[:,-2] #fixing last value for the reduced load cases
    costs_iter[scen] = pd.Series(total_costs)
    curts_iter[scen] = pd.Series(curts)
    #curts_iter.dropna(subset=scen, inplace=True)
    e_sav_iters[scen] = pd.Series(e_savs)
    e_sav_iters.dropna(subset=scen, inplace=True)
    emissions_iter[scen] = pd.Series(emissions)
    emissions_iter.dropna(subset=scen, inplace=True)
    load_curve.index = general.index[0:1440]

    iters_list = np.delete(iters_list, np.arange(len(analysis_table),len(iters_list))).tolist()
    load_curve.columns = iters_list
    analysis_table.index = iters_list
    analysis_table.dropna(axis=0, subset='Wind Output (MWh)', inplace=True)
    
    desired_range = np.arange(614,686,1).tolist()
    load_diff = pd.concat([load_curve.iloc[desired_range,0], load_curve.iloc[desired_range,-1]], axis=1)
    fig, ax = plt.subplots(figsize=(15, 3))
    duration = pd.to_datetime(load_diff.index, format='%Y%m%d.0')
    iter_n = load_diff.iloc[:,-1].values
    iter_0 = load_diff.iloc[:,0].values
    ax.plot(duration, iter_n, color='green')
    ax.plot(duration, iter_0, color='red') 
    ax.fill_between(duration, iter_n, iter_0, where=(iter_n<iter_0), interpolate=True, color="#7bad84", alpha=0.5)
    ax.fill_between(duration, iter_n, iter_0, where=(iter_n>=iter_0), interpolate=True, color="#d9db8f", alpha=0.5)
    ax.legend(handles=[Line2D([], [], color="green", label=analysis_table.index[0]), Line2D([], [], color='red', label=analysis_table.index[-1])
                      , Line2D([0], [0], color="#7bad84", label='Load Reduction', lw=6, alpha=0.5), Line2D([0], [0], color="#d9db8f", label='Load Increase', lw=6, alpha=0.5)]
                      , loc='lower center', bbox_to_anchor=(0.5, -0.35), ncol=4)    
    plt.xlabel('Time (day h)')
    #plt.xticks(desired_range, rotation=45)
    plt.ylabel('Load (MW)')
    plt.savefig(f'load_curve_changes_{scen}.jpg', dpi=300, bbox_inches='tight')
    plt.close()

    plt.figure(figsize=(15,4))
    load_curve_iter = sns.lineplot(data = load_curve.iloc[desired_range])
    load_curve_iter.set(xlabel= "Time (day h)", ylabel= "Load (MW)")
    load_curve_iter.axes.get_legend().remove()
    handles1, labels1 = load_curve_iter.axes.get_legend_handles_labels()
    plt.ylim(7000, load_curve_iter.axes.get_ylim()[1])
    wind_plt = sns.lineplot(data = wind_avail.iloc[desired_range], ax=load_curve_iter.axes.twinx(), color='#84a182', alpha=0.5, label='Available Wind')
    plt.fill_between(wind_avail.iloc[desired_range].index, wind_avail.iloc[desired_range].values, load_curve.iloc[desired_range,-1].values
                     , where=(wind_avail.iloc[desired_range].values>load_curve.iloc[desired_range,-1].values), interpolate=True, color='#ebfa6b', alpha=0.2)
    plt.fill_between(wind_avail.iloc[desired_range].index, wind_avail.iloc[desired_range].values, alpha=0.1, color='#84a182')
    handles2, labels2 = wind_plt.axes.get_legend_handles_labels()
    #handles2, labels1 = [Line2D([0], [0], color="#84a182", lw=6, alpha=0.2)], ['Available Wind']
    plt.ylim(7000, load_curve_iter.axes.get_ylim()[1])
    plt.ylabel('Available Wind (MW)')
    plt.legend(handles1+handles2, labels1+labels2, loc='lower center', bbox_to_anchor=(0.5, -0.27), ncol=len(iters_list)+1)
    plt.savefig(f'load_curves_{scen}.jpg', dpi=300, bbox_inches= 'tight')
    plt.close()

    wind_avail_cust = wind_avail.iloc[desired_range]
    fig, ax1 = plt.subplots(figsize=(11, 5))
    wind_vs_load = sns.lineplot(data=load_curve.iloc[desired_range,-1], color='g', ax=ax1)
    wind_vs_load.set(xlabel='Time (h)', ylabel='Load (MW)')
    plt.ylim(bottom=load_curve.iloc[desired_range,-1].min()-500, top=wind_avail_cust.max()+500)
    ax2 = ax1.twinx()
    sns.lineplot(data=wind_avail_cust, color='b', alpha = 0.5, ax=ax2)
    plt.ylabel('Total Available Wind (MW)')
    plt.ylim(bottom=load_curve.iloc[desired_range,-1].min()-500, top=wind_avail_cust.max()+500)
    wind_vs_load.legend(handles=[Line2D([], [], marker='_', color="g", label='Load (MW)'), Line2D([], [], marker='_', color="b", label='Total Available Wind (MW)')])
    plt.savefig(f'wind_vs_load_{scen}.jpg', dpi=300, bbox_inches='tight')
    plt.close()

    plt.figure(figsize=(8,5))
    knw_curts = curts_iter[scen].dropna()
    knw_costs = costs_iter[scen].dropna()
    x=np.arange(1,len(knw_curts)+1)
    crt_vs = sns.lineplot(x=x, y=knw_curts, color='g', marker='o')
    plt.ylim(bottom=curts_iter[scen].min()*0.98, top=curts_iter[scen].max()*1.02)
    crt_vs.set(xlabel='Iterations', ylabel='Curtailed Wind Generation (%)')
    sns.lineplot(x=x, y=knw_costs, color='b', ax=crt_vs.axes.twinx(), marker='o')
    plt.ylim(bottom=costs_iter[scen].min()*0.99, top=costs_iter[scen].max()*1.01)
    plt.ylabel('Yearly Operational Costs (M$)')
    crt_vs.legend(handles=[Line2D([], [], marker='_', color="g", label='Curtailment'), Line2D([], [], marker='_', color="b", label='Operational Costs')])
    plt.savefig(f'cost_vs_curt_{scen}.jpg', dpi=300, bbox_inches='tight')
    plt.close()

    ld_vs = sns.lineplot(data=curts_iter[scen], color='g', marker='o')
    plt.ylim(bottom=curts_iter[scen].min()*0.9, top=curts_iter[scen].max()*1.1)
    sns.lineplot(data=e_sav_iters[scen], color='r', ax=ld_vs.axes.twinx(), marker='o')
    ld_vs.legend(handles=[Line2D([], [], marker='_', color="g", label='Curtailment'), Line2D([], [], marker='_', color="r", label='Loads')])
    plt.savefig(f'loads_vs_curt_{scen}.jpg', dpi=300, bbox_inches='tight')
    plt.close()


    plt.figure(figsize=(8,5))
    crt_diff = pd.concat([curts_iter.head(1), curts_iter.tail(1)], axis=0)
    cst_diff = pd.concat([costs_iter.head(1), costs_iter.tail(1)], axis=0)
    crt_vs = sns.lineplot(data=crt_diff[scen], color='g', marker='o')
    sns.lineplot(data=cst_diff[scen], color='b', ax=crt_vs.axes.twinx(), marker='o')
    crt_vs.legend(handles=[Line2D([], [], marker='_', color="g", label='Curtailment'), Line2D([], [], marker='_', color="b", label='Operational Costs')])
    plt.savefig(f'final_cost_vs_curt_{scen}.jpg', dpi=300, bbox_inches='tight')
    plt.close()

    #plot_title = f"Curtailed Wind in % - Scenario: {row['elec_scen']}, {row['emit_scen']} - DR Method: {row['dr_method']} with {row['load_chg']}% participation"

    
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
  
    del curts
    del load_curve
    del e_savs
    del total_costs
    del emissions

    analysis_table.iloc[-1, :] = analysis_table.iloc[-2, :].astype('float64')
    last_iter = analysis_table.iloc[-1,:]
    analysis_table.loc['Lowest Cost Diff'] = analysis_table.loc[analysis_table['Operational Costs (M$)'].idxmin()]-analysis_table.iloc[0]
    analysis_table.loc['Diffrence'] = last_iter-analysis_table.iloc[0]
    analysis_table.loc['% Diffrence'] = (last_iter-analysis_table.iloc[0])/analysis_table.iloc[0]
    analysis_table.loc['Best Curtailment Diff'] = analysis_table.loc[analysis_table['Wind Output (MWh)'].idxmax()]-analysis_table.iloc[0]
    writer = pd.ExcelWriter(f'analysis_table_{scen}.xlsx', engine='xlsxwriter')
    analysis_table.to_excel(writer, sheet_name='Analysis Table')
    costs_iter.to_excel(writer, sheet_name='Costs')
    curts_iter.to_excel(writer, sheet_name='Curtailments')
    writer.save()

    os.chdir(f'{pwd}/{dir}')

    analysis_scen.loc[scen,'Savings (M$)'] = -analysis_table.loc['Lowest Cost Diff', 'Operational Costs (M$)']
    analysis_scen.loc[scen,'Dispatched Wind (MWh)'] = analysis_table.loc['Best Curtailment Diff', 'Wind Output (MWh)']
    analysis_scen.loc[scen,'Emission Reduction (Mton CO2eq)'] = analysis_table.loc['Lowest Cost Diff', 'Emissions (MTCO2e)']
    analysis_scen.loc[scen,'Savings in Household Energy Costs (M$)'] = -analysis_table.loc['Lowest Cost Diff', 'Household Energy Expenses (M$)']

    del analysis_table
    del iters_list

    

if len(scen_num) > 1:

    plt.figure(figsize=(8,5))
    crt_total = sns.lineplot(data = curts_iter)
    crt_total.set(xlabel= "Iterations", ylabel= "Curtailed Wind genartion (%)")
    plt.ylim(bottom=curts_iter[scen].min()*0.9, top=curts_iter[scen].min()*1.1)
    crt_total.axes.get_legend().remove()
    sns.lineplot(data=costs_iter, ax=crt_total.axes.twinx())
    plt.ylabel('Total Operational Costs (M$)')
    plt.ylim(bottom=costs_iter[scen].min()*0.9, top=costs_iter[scen].min()*1.1)
    plt.savefig(f"curtailments_and_costs_{scens.iat[scen_num[-1],1]}_{scens.iat[scen_num[-1],2]}_{scens.iat[scen_num[-1],4]}.jpg", dpi=300, bbox_inches= 'tight')
    plt.close()

    #Manual fix for wind and costs plot for ELEC_zeroemit scenario
    #analysis_scen['Savings (M$)'] = [17.85 ,26.62, 36.8699999999999, 40.3000000000002]
    #analysis_scen.at['ELEC_zeroemit_15%particip_DLC', 'Dispatched Wind (MWh)'] = 386946.938791633
    
    all_scens = scens.loc[scen_num]
    all_scens.index = all_scens['scen']
    analysis_scen['Participation (%)'] = all_scens['particip']
    cost_and_winds = sns.lineplot(x='Participation (%)', y='Savings (M$)', data=analysis_scen, marker='o', color='b')
    plt.ylabel('Operational Costs Savings (M$)')
    plt.ylim(bottom=0.9*analysis_scen['Savings (M$)'].min(), top=1.1*analysis_scen['Savings (M$)'].max())
    sns.lineplot(x='Participation (%)', y='Dispatched Wind (MWh)', data=analysis_scen, marker='o', ax=cost_and_winds.axes.twinx(), color='r')
    plt.ylabel('Improved Wind Integration (MWh)')
    plt.ylim(bottom=0.9*analysis_scen['Dispatched Wind (MWh)'].min(), top=1.1*analysis_scen['Dispatched Wind (MWh)'].max())
    plt.xticks(analysis_scen['Participation (%)'])
    cost_and_winds.legend(handles=[Line2D([], [], marker='_', color="b", label='Operational Costs Savings (M$)'), Line2D([], [], marker='_', color="r", label='Improved Wind Integration (MWh)')])
    plt.savefig(f"costs_and_winds_{scens.iat[scen_num[-1],1]}_{scens.iat[scen_num[-1],2]}_{scens.iat[scen_num[-1],4]}.jpg", dpi=300, bbox_inches= 'tight')
    plt.close()

    esav_total = sns.lineplot(data = e_sav_iters)
    esav_total.set(xlabel= "Iterations", ylabel= "Total Energy Supplied (MWh)")
    plt.savefig(f"esavings_{scens.iat[scen_num[-1],1]}_{scens.iat[scen_num[-1],2]}_{scens.iat[scen_num[-1],4]}.jpg", dpi=300, bbox_inches= 'tight')
    plt.close()

    analysis_scen.to_excel(f"analysis_scen_{scens.iat[scen_num[-1],1]}_{scens.iat[scen_num[-1],2]}_{scens.iat[scen_num[-1],4]}.xlsx")
