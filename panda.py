import pandas as pd

# d = pd.read_csv('data.csv')

# print(d.to_string())



# print(pd.options.display.max_rows)

fp= 'new.xlsx'

d = pd.read_excel(fp)

'''
# new_file= d.dropna()
# print(new_file)

# d.dropna(inplace=True)
# print(d)

##Fill empty values

# d.fillna(100,inplace=True)
# print(d)
'''

group = d.groupby('STUDIED')

for STUDIED, g in group:
    f_name = f"{STUDIED}.xlsx"
    g.to_excel(f_name,index=False)
    print(f"Saved {f_name}")