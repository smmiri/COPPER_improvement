import csv
import os
import pandas as pd

def fixpath(path):
    if path.startswith("C:"): return "/mnt/c/" + path.replace("\\", "/")[3:]
    else:
        pass
    return path

os.chdir(fixpath(r'C:\Users\smoha\Documents\git\copper\COPPER7'))

coordinate_df = pd.read_excel("coordinate.xlsx", sheet_name="coordinate_system")

coordinate_BC = coordinate_df.loc[coordinate_df['PRENAME'] == 'British Columbia']
coordinate_AB = coordinate_df.loc[coordinate_df['PRENAME'] == 'Alberta']



with open('windcf.csv') as csv_file:
    reader = csv.reader(csv_file)
    windcf = dict(reader)

for k in windcf:
    windcf[k]=float(windcf[k])

with open('solarcf.csv') as csv_file:
    reader = csv.reader(csv_file)
    solarcf = dict(reader)

for k in solarcf:
    solarcf[k]=float(solarcf[k])



windcf_BC = [windcf[i] for i in map(list(str), f'*.{coordinate_BC.loc[:,"grid cell"]}')]
print(windcf_BC)