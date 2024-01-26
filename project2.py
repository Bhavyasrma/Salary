import pandas as pd
import xlwings as xw
import sys
from library import get_index_nm
from library import get_index_normal
from library import bkt_name



def Salary_sheet(arg1):
    PATH = arg1
    wb = xw.Book(PATH)
    sheet = wb.sheets['DATA']

    df = sheet['B:M'].options(pd.DataFrame, index=False, header=True).value
    df = df.dropna(how="all")
    entities = df.FOS.unique()
    products = ["PL SAL", "PL SELF", "DIGITAL"]
    bkt_types = ["FE", 1, 2, 3,4]
    column_sum={}
    column_sum2={}
    enity_sum={}
    enity_sum2 = {}
    dfs = {}
    dfs2 = {}
    dfs3 = {}
    dfs4 = {}
    for entity in entities:
        column_sum ={}
        column_sum2 = {}
        for bkt in bkt_types:
            for product in products:
                key = f"df_{entity}_{product}_{bkt}"
                dfs[key] = df.loc[(df['FOS'] == entity) & (df['PRODUCT'] == product) & (df['BKT'] == bkt) & (df['Status'] == 'PAID')]
                dfs2[key] = df.loc[(df['FOS'] == entity) & (df['PRODUCT'] == product) & (df['BKT'] == bkt) & (df['Status'] == 'UNPAID')]
                dfs3[key] = df.loc[(df['FOS'] == entity) & (df['PRODUCT'] == product) & (df['BKT'] == bkt) & (df['RB'] == 'NM')]
                dfs4[key] = df.loc[(df['FOS'] == entity) & (df['PRODUCT'] == product) & (df['BKT'] == bkt)]
                condition = dfs[key]['POS'].sum()
                condition2 = dfs2[key]['POS'].sum()
                condition3 = dfs3[key]['POS'].sum()
                condition4 = dfs4[key]['POS'].sum()
                total = condition + condition2
                result = (condition/total) if total != 0 else 0
                nm_result = (condition3/condition4) if total != 0 else 0           
                column_sum[key]=result
                column_sum2[key]= nm_result
        enity_sum[entity]  = column_sum
        enity_sum2[entity]  = column_sum2
    sheet2 = wb.sheets['Payout']
    df2 = sheet2['A:I'].options(pd.DataFrame, index=False, header=True).value
    df2 = df2.dropna(how="all")

    df_FE = df2.iloc[1:3 ,0:6]
    df_FE.columns = df_FE.iloc[0]
    df_FE = df_FE[1:]
    df_FE.set_index('NM', inplace=True)

    df_bkt1_PLSAL = df2.iloc[4:10 ,0:6]
    df_bkt1_PLSAL.columns = df_bkt1_PLSAL.iloc[0]
    df_bkt1_PLSAL = df_bkt1_PLSAL[1:]
    df_bkt1_PLSAL.set_index('NM', inplace=True)

    df_bkt1_PLSELF = df2.iloc[11:16 ,0:6]
    df_bkt1_PLSELF.columns = df_bkt1_PLSELF.iloc[0]
    df_bkt1_PLSELF = df_bkt1_PLSELF[1:]
    df_bkt1_PLSELF.set_index('NM', inplace=True)

    df_bkt2 = df2.iloc[17:22 ,0:5] 
    df_bkt2.columns = df_bkt2.iloc[0]
    df_bkt2 = df_bkt2[1:]
    df_bkt2.set_index('NM', inplace=True)


    df_bkt4 = df2.iloc[23:28 ,0:5]
    df_bkt4.columns = df_bkt4.iloc[0]
    df_bkt4 = df_bkt4[1:]
    df_bkt4.set_index('NM', inplace=True)


    anss = {}
    Dict1 = {'FE': df_FE, '1_PL SAL': df_bkt1_PLSAL, '1_PL SELF':df_bkt1_PLSELF,'1_DIGITAL':df_bkt1_PLSAL, '2':df_bkt2, '3':df_bkt2, '4':df_bkt4}
    entities = df.FOS.unique()
    products = ["PL SAL", "PL SELF", "DIGITAL"]
    bkt_types = ["FE", 1, 2, 3,4]
    column_sum={}
    entity_sum={}
    dfs = {}
    for entity in entities:
        column_sum ={}
        for bkt in bkt_types:
            for product in products:
                key = f"df_{entity}_{product}_{bkt}"
                dfs[key] = df.loc[(df['FOS'] == entity) & (df['PRODUCT'] == product) & (df['BKT'] == bkt)]
                condition = dfs[key]['ACTUAL EMI'].sum()
                column_sum[key]=condition
        entity_sum[entity]  = column_sum      
    
    writer = pd.ExcelWriter("Salary_sheet.xlsx", engine="xlsxwriter")
    workbook = writer.book
    worksheet = workbook.add_worksheet()
    merge_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1,'bold' : True})
    worksheet.merge_range(0, 0, 0, 15, 'DEC23', merge_format)
    worksheet.merge_range(1, 0, 2, 0, 'Name', merge_format)

    worksheet.merge_range(1, 1, 1, 3, 'FR', merge_format)
    worksheet.merge_range(1, 4, 1, 6, 'BKT1', merge_format)
    worksheet.merge_range(1, 7, 1, 9, 'BKT2', merge_format)
    worksheet.merge_range(1, 10, 1, 12, 'BKT3', merge_format)
    worksheet.merge_range(1, 13, 1, 15, 'BKT4', merge_format)

    categories = ['SAL', 'SELF', 'DIGITAL']
    start_col = 0

    for j in range(5):
        for  category in categories:
            start_col = start_col + 1
            worksheet.write(2, start_col, category, merge_format)
            
        
    start_row = 3  
    for j, key in enumerate(entity_sum):
        worksheet.write(start_row+j, 0, key, merge_format)
        for i, key_sum in enumerate(entity_sum[key]):
          worksheet.write(3+j, 1+i, entity_sum[key][key_sum])

    worksheet.merge_range(13, 0, 13, 12, 'DEC23', merge_format)
    worksheet.merge_range(14, 0, 15, 1, 'Prod. Name', merge_format)
    worksheet.merge_range(14, 2, 14, 12, 'NAME', merge_format)

    start_column =  2
    for j, key in enumerate(entity_sum):
        worksheet.write(15, start_column+j, key, merge_format)

    categories = ['SAL', 'SELF', 'DIGITAL']
    start_col = 15

    for j in range(5):
        for  category in categories:
            start_col = start_col + 1
            worksheet.write(start_col, 1, category, merge_format)

    #payout indexing and final result

    for i,key in enumerate(enity_sum):
        for j,key2 in enumerate(enity_sum[key]):
            v1 = enity_sum[key][key2]
            v1 = round(v1, 2)
            v2 = enity_sum2[key][key2]
            v2 = round(v2, 2)
            bkt_name_final = bkt_name(key2)
            index1 = get_index_normal(Dict1[bkt_name_final],v1)
            index2 =0
            index2 = get_index_nm(Dict1[bkt_name_final],v2)
            value = Dict1[bkt_name_final]
            val = value.loc[index2,index1]
            pay = entity_sum[key][key2]
            ans = val*pay
            anss[key2]= ans
            
            bkt_split = bkt_name_final.split('_')
            bkt_split = bkt_split[0]
            worksheet.write(16+j, 0, bkt_split, merge_format)
            
            worksheet.write(16+j, 2+i, anss[key2])



    writer.save()
    writer.close()
          

  

if __name__ == "__main__" :
    arg1 = sys.argv[1]
    Salary_sheet(arg1)

