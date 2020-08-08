import pandas as pd
import stdnum as sd
import yfinance as yf

d = pd.read_excel('GICS_map.xlsx', header = 4)
d = d.rename(columns = {'Unnamed: 1': 'SectorDesc'})
d = d.rename(columns = {'Unnamed: 3': 'GroupDesc'})
d = d.rename(columns = {'Unnamed: 5': 'IndustrypDesc'})
d = d.rename(columns = {'Unnamed: 7': 'SubIndustrypDesc'})

d = d.ffill(axis = 0)

r = pd.read_excel('Russell 3000 - Subindustry.xlsx', header = 0)
r = r.rename(columns = {'Unnamed: 0': 'SubIndustry'})

r['SubIndustry'] = r['SubIndustry'].ffill(axis = 0)
r.drop(r[r['Russell 3000'].isna()].index, inplace = True)
r.drop(r[r['SEDOL'].isna()].index, inplace = True)

sedol = get_cc_module('gb', 'sedol')
r['isin'] = r.SEDOL.apply(sedol.to_isin)
