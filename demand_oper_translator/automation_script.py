from glob import glob
import os
from Automate_OS_and_Measures_fixed import print_to_osw, run_from_cl
#print_to_osw('ZEBRA', 'GIRAFFE', 'HIPPO')

cwd = os.getcwd()
pwd = os.path.dirname(cwd)
os.chdir(pwd)
directories = glob('archetypes_base/*/', recursive=True)
naming = input("Input the iteration number province, and scenario in the correct format:\n(iteration_province_scenario in the format, iteration_elecscenario_emissionscenario)\n")

house_types = ['Apartment Shape_iter1_ELEC_emit', 'Apartment Shape', 'new house shape', 'old house shape']

'''for dir in directories:

    house_type = f'{dir[16:-1]}_{naming}'
    house_types.append(house_type)

house_types.insert(0, 'Apartment Shape')'''


os.chdir(cwd)

for htype in house_types:

    print(htype)
    print_to_osw(htype, naming, 'Heating:Electricity')
    run_from_cl('WorkFlow')
