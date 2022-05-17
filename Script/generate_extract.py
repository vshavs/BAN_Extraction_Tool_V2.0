import pandas as pd
import itertools


def data_extract(file, ban1, ban2, path1, path2, env1, env2):
    ensemble_df = pd.read_excel(file, sheet_name='Ensemble Query', index_col=None)
    magenta_df = pd.read_excel(file, sheet_name='Magenta Query', index_col=None)
    data1 = []
    data2 = []
    data3 = []
    data4 = []
    data5 = []
    data6 = []
    for i, row in ensemble_df.iterrows():
        var1 = row['IND']
        data1.append(var1)
        var2 = row['Table_Name']
        data2.append(var2)
        var3 = row['Queries'].format(ban1)
        data3.append(var3)
    for j, row in magenta_df.iterrows():
        var4 = row['IND']
        data4.append(var4)
        var5 = row['Table_Name']
        data5.append(var5)
        var6 = row['Queries'].format(ban2)
        data6.append(var6)
    with pd.ExcelWriter(path1, engine='xlsxwriter', date_format='%m/%d/%Y_%H:%M:%S') as ens_writer :
        for i, j, k in itertools.zip_longest(data1, data2, data3):
            if i == "Y":
                ensemble_data = pd.read_sql(k, env1)
                ensemble_data.to_excel(ens_writer, index=False, sheet_name=j)
        print("Ensemble extract is generated")
    with pd.ExcelWriter(path2, engine='xlsxwriter', date_format='%m/%d/%Y_%H:%M:%S') as mag_writer:
        for l, m, n in itertools.zip_longest(data4, data5, data6):
            if l == "Y":
                magenta_data = pd.read_sql(n, env2)
                magenta_data.to_excel(mag_writer, index=False, sheet_name=m)
        print("Metro extract is generated")
