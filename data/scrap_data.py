import pyilt2
import pyilt2.report
from tabulate import tabulate
from collections import OrderedDict
import xlsxwriter
import re
import pandas as pd
import numpy as np

results = pyilt2.query(numOfComp = 2, prop = pyilt2.prop2abr["Henry's Law constant"])

results_list = pyilt2.report.getAllData(results)
Hcs, temps, solutes = [], [], []
for result in results_list:
  data = result.data
  header = result.headerList

  # convert data and header to dataframe
  df = pd.DataFrame(data, columns=header)

  # find column whose name starts with Henry's_Law_constant and rename it to Hc
  hc_col = [col for col in df.columns if col.startswith('Henry\'s_Law_constant')][0]
  df.rename(columns={hc_col: 'Hc'}, inplace=True)

  # find column whose name starts with Temperature and rename it to temp
  temp_col = [col for col in df.columns if col.startswith('Temperature')][0]
  df.rename(columns={temp_col: 'temp'}, inplace=True)

  # find column whose name starts with Mole_fraction_of_ and store the name in variable
  mol_frac_col = [col for col in df.columns if '[Liquid]' in col][0]
  # from mol_frac_col extract the part of the string between ' Mole_fraction_of_' and '[Liquid]'
  mol_frac_col = mol_frac_col.split('_of_')[1].split('[Liquid]')[0]
  if '_(normal)' in mol_frac_col:
    mol_frac_col = mol_frac_col.split('_(normal)')[0]

  # add to df column 'solute' and store in it mol_frac_col for all rows in df
  df['solute'] = mol_frac_col

  # remove all other columns
  df = df[['Hc', 'temp', 'solute']]
  Hcs.extend(df['Hc'].values)
  temps.extend(df['temp'].values)
  solutes.extend(df['solute'].values)
df = pd.DataFrame()
df['Hcs'] = Hcs
df['temps'] = temps
df['solutes'] = solutes
solutes_to_select = ['hydrogen', 'carbon_dioxide', 'oxygen', 'nitrogen', 'methane', 'ethane', 'argon', 'carbon_monoxide', 'ethene', 'propane', 'propene', 'ammonia', 'hydrogen_sulfide', 'water', 'pentafluoroethane', 'difluoromethane', 'trifluoromethane', 'sulfur_dioxide', 'xenon', '2-methylpropane', 'butane', '1-butene', 'krypton']
df = df[df['solutes'].isin(solutes_to_select)]
df.reset_index(drop=True, inplace=True)
df.to_excel("henry.xlsx")