import pandas as pd
import os

ctax = [0, 50, 100, 150, 200]
ap = ["British Columbia", "Alberta", "Saskatchewan", "Manitoba", "Ontario", "Quebec", "New Brunswick",
      "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island"]  # all provinces
pds = ['2030', '2040', '2050']
gendata = pd.read_excel(r'Generation_type_data_SMR_CCS.xlsx', header=0)
allplants = list(gendata.iloc[:]['Type'])
ramp_rate_percent = dict(zip(list(gendata.iloc[:]['Type']), list(gendata.iloc[:][
                                                                     'ramp_rate_percent'])))  # {'gas': 0.1, 'peaker':0.1 , 'nuclear': 0.05,  'coal': 0.05, 'diesel': 0.1, 'waste': 0.05} #ramp rate in percent of capacity per hour

cwd = os.getcwd()

for i in ctax:

    os.chdir(cwd + '\outputs_ct' + str(i) + '_rd38_pds3_Hr_OBPS_LGP_NoHydro_NoCL_CPHy_NoAr_SMR_CCS_CPO_GPS')

    cap_ap = pd.read_excel(r'Total_generation_ap.xlsx', header=0, index_col=0)

    # print(cap_ap.loc['coal','British Columbia'])

    cap_ap_2030 = cap_ap.iloc[[0, 1, 2, 3, 4, 5, 6, 7, 8, 11], :]
    cap_ap_2040 = cap_ap.iloc[[12, 13, 14, 15, 16, 17, 18, 19, 20, 23], :]
    cap_ap_2050 = cap_ap.iloc[[24, 25, 26, 27, 28, 29, 30, 31, 32, 35], :]

    # print(cap_ap_2030,cap_ap_2040,cap_ap_2050)

    cap_ap_2030_total = cap_ap_2030.agg(['sum'])
    cap_ap_2030_total = cap_ap_2030_total.append([cap_ap_2030_total] * 11)
    cap_ap_2030_total.index = allplants
    cap_ap_2040_total = cap_ap_2040.agg(['sum'])
    cap_ap_2040_total = cap_ap_2040_total.append([cap_ap_2040_total] * 11)
    cap_ap_2040_total.index = allplants
    cap_ap_2050_total = cap_ap_2050.agg(['sum'])
    cap_ap_2050_total = cap_ap_2050_total.append([cap_ap_2050_total] * 11)
    cap_ap_2050_total.index = allplants

    # print(cap_ap_2030 / cap_ap_2030_total)
    # cap_total = cap_ap.agg()
    # print(cap_total)

    flex_i_2030 = pd.DataFrame()
    flex_i_2040 = pd.DataFrame()
    flex_i_2050 = pd.DataFrame()

    # print(cap_ap_2030_total)

    for ALP in allplants:
        if ALP == 'coal' or ALP == 'coalccs' or ALP == 'diesel':
            flex_i_2030 = flex_i_2030.append(
                (cap_ap_2030.loc[ALP, :] / cap_ap_2030_total.loc[ALP, :]) * ((0.5 * (cap_ap_2030.loc[ALP, :])) + 0.5 * (
                        cap_ap_2030.loc[ALP, :] * ramp_rate_percent[ALP])) / cap_ap_2030.loc[ALP, :])
            flex_i_2040 = flex_i_2040.append(
                (cap_ap_2040.loc[ALP, :] / cap_ap_2040_total.loc[ALP, :]) * ((0.5 * (cap_ap_2040.loc[ALP, :])) + 0.5 * (
                        cap_ap_2040.loc[ALP, :] * ramp_rate_percent[ALP])) / cap_ap_2040.loc[ALP, :])
            flex_i_2050 = flex_i_2050.append(
                (cap_ap_2050.loc[ALP, :] / cap_ap_2050_total.loc[ALP, :]) * ((0.5 * (cap_ap_2050.loc[ALP, :])) + 0.5 * (
                        cap_ap_2050.loc[ALP, :] * ramp_rate_percent[ALP])) / cap_ap_2050.loc[ALP, :])
        elif ALP == 'gas' or ALP == 'gasccs':
            flex_i_2030 = flex_i_2030.append(
                (cap_ap_2030.loc[ALP, :] / cap_ap_2030_total.loc[ALP, :]) * (
                            (0.25 * (cap_ap_2030.loc[ALP, :])) + 0.5 * (
                            cap_ap_2030.loc[ALP, :] * ramp_rate_percent[ALP])) / cap_ap_2030.loc[ALP, :])
            flex_i_2040 = flex_i_2040.append(
                (cap_ap_2040.loc[ALP, :] / cap_ap_2040_total.loc[ALP, :]) * (
                            (0.25 * (cap_ap_2040.loc[ALP, :])) + 0.5 * (
                            cap_ap_2040.loc[ALP, :] * ramp_rate_percent[ALP])) / cap_ap_2040.loc[ALP, :])
            flex_i_2050 = flex_i_2050.append(
                (cap_ap_2050.loc[ALP, :] / cap_ap_2050_total.loc[ALP, :]) * (
                            (0.25 * (cap_ap_2050.loc[ALP, :])) + 0.5 * (
                            cap_ap_2050.loc[ALP, :] * ramp_rate_percent[ALP])) / cap_ap_2050.loc[ALP, :])
        elif  ALP == 'peaker':
            flex_i_2030 = flex_i_2030.append(
                (cap_ap_2030.loc[ALP, :] / cap_ap_2030_total.loc[ALP, :]) * (
                            ((cap_ap_2030.loc[ALP, :])) + 0.5 * (
                            cap_ap_2030.loc[ALP, :] * ramp_rate_percent[ALP])) / cap_ap_2030.loc[ALP, :])
            flex_i_2040 = flex_i_2040.append(
                (cap_ap_2040.loc[ALP, :] / cap_ap_2040_total.loc[ALP, :]) * (
                            ((cap_ap_2040.loc[ALP, :])) + 0.5 * (
                            cap_ap_2040.loc[ALP, :] * ramp_rate_percent[ALP])) / cap_ap_2040.loc[ALP, :])
            flex_i_2050 = flex_i_2050.append(
                (cap_ap_2050.loc[ALP, :] / cap_ap_2050_total.loc[ALP, :]) * (
                            ((cap_ap_2050.loc[ALP, :])) + 0.5 * (
                            cap_ap_2050.loc[ALP, :] * ramp_rate_percent[ALP])) / cap_ap_2050.loc[ALP, :])
        elif ALP == 'nuclear':
            flex_i_2030 = flex_i_2030.append(
                (cap_ap_2030.loc[ALP, :] / cap_ap_2030_total.loc[ALP, :]) * ((0.5 * (cap_ap_2030.loc[ALP, :])) + 0.5 * (
                        cap_ap_2030.loc[ALP, :] * ramp_rate_percent[ALP])) / cap_ap_2030.loc[ALP, :])
            flex_i_2040 = flex_i_2040.append(
                (cap_ap_2040.loc[ALP, :] / cap_ap_2040_total.loc[ALP, :]) * ((0.5 * (cap_ap_2040.loc[ALP, :])) + 0.5 * (
                        cap_ap_2040.loc[ALP, :] * ramp_rate_percent[ALP])) / cap_ap_2040.loc[ALP, :])
            flex_i_2050 = flex_i_2050.append(
                (cap_ap_2050.loc[ALP, :] / cap_ap_2050_total.loc[ALP, :]) * ((0.5 * (cap_ap_2050.loc[ALP, :])) + 0.5 * (
                        cap_ap_2050.loc[ALP, :] * ramp_rate_percent[ALP])) / cap_ap_2050.loc[ALP, :])
        elif ALP == 'waste':
            flex_i_2030 = flex_i_2030.append(
                (cap_ap_2030.loc[ALP, :] / cap_ap_2030_total.loc[ALP, :]) * ((0.2 * (cap_ap_2030.loc[ALP, :])) + 0.5 * (
                        cap_ap_2030.loc[ALP, :] * ramp_rate_percent[ALP])) / cap_ap_2030.loc[ALP, :])
            flex_i_2040 = flex_i_2040.append(
                (cap_ap_2040.loc[ALP, :] / cap_ap_2040_total.loc[ALP, :]) * ((0.2 * (cap_ap_2040.loc[ALP, :])) + 0.5 * (
                        cap_ap_2030.loc[ALP, :] * ramp_rate_percent[ALP])) / cap_ap_2040.loc[ALP, :])
            flex_i_2050 = flex_i_2050.append(
                (cap_ap_2050.loc[ALP, :] / cap_ap_2050_total.loc[ALP, :]) * ((0.2 * (cap_ap_2050.loc[ALP, :])) + 0.5 * (
                        cap_ap_2030.loc[ALP, :] * ramp_rate_percent[ALP])) / cap_ap_2050.loc[ALP, :])
        elif ALP == 'hydro':
            flex_i_2030 = flex_i_2030.append(
                (cap_ap_2030.loc[ALP, :] / cap_ap_2030_total.loc[ALP, :]) * ((0.5 * (cap_ap_2030.loc[ALP, :])) + 0.5 * (
                        cap_ap_2030.loc[ALP, :] * ramp_rate_percent[ALP])) / cap_ap_2030.loc[ALP, :])
            flex_i_2040 = flex_i_2040.append(
                (cap_ap_2040.loc[ALP, :] / cap_ap_2040_total.loc[ALP, :]) * ((0.5 * (cap_ap_2040.loc[ALP, :])) + 0.5 * (
                        cap_ap_2040.loc[ALP, :] * ramp_rate_percent[ALP])) / cap_ap_2040.loc[ALP, :])
            flex_i_2050 = flex_i_2050.append(
                (cap_ap_2050.loc[ALP, :] / cap_ap_2050_total.loc[ALP, :]) * ((0.5 * (cap_ap_2050.loc[ALP, :])) + 0.5 * (
                        cap_ap_2050.loc[ALP, :] * ramp_rate_percent[ALP])) / cap_ap_2050.loc[ALP, :])

    total_flex = pd.concat([flex_i_2030.agg(['sum']), flex_i_2040.agg(['sum']), flex_i_2050.agg(['sum'])])
    total_flex.index = [2030, 2040, 2050]
    total_flex.to_excel('ct' + str(i) + '_total_flex_index_ap.xlsx', index=True)

    os.chdir(cwd)

    #print(cap_ap)

print(flex_i_2030, flex_i_2030.agg(['sum']))