
import pandas as pd
import os
import re
import xlsxwriter
import openpyxl
from datetime import datetime
from datetime import timedelta
from datetime import date
import numpy as np
import webbrowser

os.getcwd()

print('PROCESSING DIRECTORY...')
getmonth = str.upper(datetime.now().strftime("%b"))
getyear = datetime.now().strftime("%Y")
name_dir = "working_files_{month} {year}"
directory = name_dir.format(year=getyear, month=getmonth)
parent_dir = "*"
dir_path = os.path.join(parent_dir, directory)
isExist = os.path.exists(dir_path)

if not isExist:
    os.makedirs(dir_path)
    print("Directory '%s' created" %directory)
else:
    print("Directory '%s' exists" %directory)

print('#'*28)

productcatalogexpanded = str(input('Paste path productCatalogExpanded here: ')).replace('"','')
tech_info = pd.read_excel(productcatalogexpanded,
                         dtype={'ECLASS':str})

print('File Currency is: ', set(tech_info['CURR']))
if 'Dollar' in set(tech_info['CURR']):
    curr = 'USD'
elif 'Euro' in set(tech_info['CURR']):
    curr = 'EUR'
elif 'GBPound' in set(tech_info['CURR']):
    curr = 'GBP'
else:
    curr = 'SEK'

curr_productcatalogexpanded = "{}\{}productCatalogExpanded.xls".format(dir_path, curr)
os.rename(productcatalogexpanded, curr_productcatalogexpanded)
print('File is renamed.')
print(tech_info.isna().sum())

print('\nPROCESSING DATA CLEANING...')
tech_info['LEVEL3'] = tech_info.groupby('LEVEL2')['LEVEL3'].fillna(tech_info['LEVEL3'].mode()[0])
tech_info['UNSPSCCODE'] = tech_info.groupby('LEVEL3')['UNSPSCCODE'].fillna(tech_info['UNSPSCCODE'].mode()[0])
tech_i = tech_info[['PARTNUM', 'PTITLE', 'QTYAMOUNT', 'VOLUMEPRICE', 'PRODUCTIMAGE', 'PAGELINK', 'UNITS', 'LEVEL2', 'LEVEL3', 'LEVEL4', 'PRODUCTWEIGHT', 'UNSPSCCODE', 'ECLASS', 'KEYWORDS']]
tech_i = tech_i[tech_i['VOLUMEPRICE'] != 0].reset_index(drop=True)
tech_i = tech_i[tech_i['QTYAMOUNT'] ==1].reset_index(drop=True)
cwf = tech_i.drop_duplicates(subset=['PARTNUM']).reset_index(drop=True)
cwf['SUPPLIER_AID'] = cwf.apply(lambda _: '', axis=1)
cwf['PREIS'] = cwf.apply(lambda _: '', axis=1)

print('\nPROCESSING UNITS')
cwf['ORDER_UNIT'] = cwf['UNITS'].str.upper()
cwf['ORDER_UNIT'].replace({'EACH':'C62', 'METER':'MTR'}, inplace=True)
cwf['CONTENT_UNIT'] = cwf['ORDER_UNIT']
cwf.loc[cwf['PTITLE'].str.contains("\d+Pack\s+ | \d+\s+pack\s+ | Pack\s+of\s+\d+ | Package\s+of\s+\d+"), 'CONTENT_UNIT'] = 'PK'
cwf.loc[cwf['PTITLE'].str.contains("\spair\s+ | Pair\s+"), 'CONTENT_UNIT'] = 'PR'
cwf['Mengeneinheit']=cwf['ORDER_UNIT'].replace(['C62'],'PCE').reset_index(drop=True)
print('Columns set: ', cwf.columns)

print('\nPROCESSING QUANTITIES...')
cwf['Menge pro Verkaufseinheit/ NO_CU_PER_OU'] = cwf.apply(lambda _: '', axis=1)
cwf.loc[cwf['PTITLE'].str.contains("Pack\s+of\s+1000 | 1000\s+Pack\s+"), 'Menge pro Verkaufseinheit/ NO_CU_PER_OU'] = 1000
cwf.loc[cwf['PTITLE'].str.contains("200\s+Pack\s+ | Pack\s+of\s+200"), 'Menge pro Verkaufseinheit/ NO_CU_PER_OU'] = 200
cwf.loc[cwf['PTITLE'].str.contains("100\s+Pack\s+ | Pack\s+of\+100"), 'Menge pro Verkaufseinheit/ NO_CU_PER_OU'] = 100
cwf.loc[cwf['PTITLE'].str.contains("50\s+Pack\s+ | Pack\s+of\s+50"), 'Menge pro Verkaufseinheit/ NO_CU_PER_OU'] = 50
cwf.loc[cwf['PTITLE'].str.contains("25\s+Pack\s+ | Pack\s+of\s+25"), 'Menge pro Verkaufseinheit/ NO_CU_PER_OU'] = 25
cwf.loc[cwf['PTITLE'].str.contains("20\s+Pack\s+ | Pack\s+of\s+20"), 'Menge pro Verkaufseinheit/ NO_CU_PER_OU'] = 20
cwf.loc[cwf['PTITLE'].str.contains("16\s+Pack\s+ | Pack\s+of\s+16"), 'Menge pro Verkaufseinheit/ NO_CU_PER_OU'] = 16
cwf.loc[cwf['PTITLE'].str.contains("10\s+Pack\s+ | Pack\s+of\s10 | 10\s+Packets"), 'Menge pro Verkaufseinheit/ NO_CU_PER_OU'] = 10
cwf.loc[cwf['PTITLE'].str.contains("6\s+Pack\s+ | Pack\s+of\s+6"), 'Menge pro Verkaufseinheit/ NO_CU_PER_OU'] = 6
cwf.loc[cwf['PTITLE'].str.contains("5\s+Pack\s+ | 5\s+Packages | Pack\s+of\s+5"), 'Menge pro Verkaufseinheit/ NO_CU_PER_OU'] = 5
cwf.loc[cwf['PTITLE'].str.contains("4\s+Pack\s+ | Pack\s+of\s+4"), 'Menge pro Verkaufseinheit/ NO_CU_PER_OU'] = 4
cwf.loc[cwf['PTITLE'].str.contains("Pack\s+of\s+2\s | 2\s+Pack\s+"), 'Menge pro Verkaufseinheit/ NO_CU_PER_OU'] = 2 #bylo 200
cwf['Menge pro Verkaufseinheit/ NO_CU_PER_OU'] = cwf['Menge pro Verkaufseinheit/ NO_CU_PER_OU'].replace([''],'1')
cwf['QUANTITY_MIN'] = '1'
cwf['PRICE_QUANTITY'] = '1'

print('\nPROCESSING PRODUCT DESCRIPTION...')
cwf['PTITLE'].replace(to_replace=['</sup>/', '</sup>', '<sup>&reg;', '<sup>', '<sup style="font-size:6pt;', '">'], 
                      value=' ', regex=True)
cwf['DESCRIPTION_SHORT (80)'] = cwf['PTITLE'].reset_index(drop=True)
cwf['DESCRIPTION_SHORT (80)'] = cwf['DESCRIPTION_SHORT (80)'].str.slice(start=0, stop=80)
cwf['DESCRIPTION_LONG'] = cwf['PTITLE']

print('\nPROCESSING CATALOG GROUP IDs...')
cwf['CATALOG_GROUP_ID'] = cwf.apply(lambda _: '', axis=1)
cgID = tech_info[['LEVEL2']].drop_duplicates(subset=['LEVEL2']).reset_index(drop=True)
cgID.head()
cgID = cgID.reset_index(drop=True)
cgID.index = cgID.index + 201
cgID['cg ID'] = cgID.index
cwf = pd.merge(cwf, cgID,
               left_on='LEVEL2', 
               right_on='LEVEL2', 
               how='left', 
               suffixes=("", "_temp")).reset_index(drop=True)
cwf['CATALOG_GROUP_ID'] = cwf['cg ID']

print('\nPROCESSING TRADE DATA...')
cwf['MANUFACTURER_NAME'] = 'Thorlabs'
cwf['MANUFACTURER_AID'] = cwf.apply(lambda _: '', axis=1)
cwf['DELIVERY_TIME'] = 2
cwf['KEYWORD (max 50)'] = cwf.apply(lambda _: '', axis=1)
level234 = tech_info[['PARTNUM', 'LEVEL2', 'LEVEL3', 'LEVEL4']].reset_index(drop=True)
level234 = level234.drop_duplicates(subset=['PARTNUM']).reset_index(drop=True)
level234['KEYWORD'] = [str(x) + ', ' + str(y) for x, y in zip(level234['LEVEL2'], level234['LEVEL3'])]
cwf = pd.merge(cwf, level234,
               left_on='PARTNUM', 
               right_on='PARTNUM', 
               how='left', 
               suffixes=("", "_temp")).reset_index(drop=True)
cwf['KEYWORD (max 50)'] = cwf['KEYWORD'].str.slice(start=0, stop=50)
cwf['KEYWORD (max 50)'] = cwf['KEYWORD (max 50)'].replace(['&amp;'],'&')

print('\nPROCESSING MIMEs...')
cwf['MIME_SOURCE1'] = cwf.apply(lambda _: '', axis=1)
cwf['MIME_SOURCE1_lrg'] = cwf.apply(lambda _: '', axis=1)
cwf['MIME_SOURCE1_lrg_weblink'] = cwf.apply(lambda _: '', axis=1)
cwf['MIME_SOURCE1_lrg_weblink'] = cwf['PRODUCTIMAGE'].replace(to_replace = ['small', 'sm'],
                                                              value=['large', 'lrg'], regex=True)

cwf['MIME_SOURCE1'] = cwf['PRODUCTIMAGE'].replace(to_replace='http://www.thorlabs.com/images/small/',
                                                  value='', regex=True)

cwf['MIME_SOURCE1_lrg'] = cwf['MIME_SOURCE1_lrg_weblink'].replace(to_replace='http://www.thorlabs.com/images/large/',
                                                             value='', regex=True)

cwf['MIME_PURPOSE1'] = 'normal'
cwf['MIME_DESCR1'] = 'Thumbnail'
cwf['MIME_TYPE1'] = 'image/jpeg'
cwf['MIME_SOURCE2'] = cwf.apply(lambda _: '', axis=1)
cwf['MIME_SOURCE2'] = cwf['PAGELINK']
cwf['MIME_PURPOSE2'] = 'detail'
cwf['MIME_TYPE2'] = 'url'
cwf['MIME_DESCR2'] = 'Web URL'

print('\nPROCESSING ECCN\nCommodity Code\neClass 5.1 / 6.0 / 6.2 / 7.0 / 7.1 / 10.0.1 /10.1')
itemmasterww = str(input('Item Master - WW path: ')).replace('"','')
trade_info = pd.read_excel(itemmasterww,
                           skiprows=1,
                           dtype={'Commodity':str, 'ECCN ':str})

trade_i = trade_info[['Item Number', 'Commodity', 'ECCN ']]
trade_i = trade_i.drop_duplicates(subset=['Item Number']).reset_index(drop=True)

cwf = pd.merge(cwf, trade_i, 
               left_on='PARTNUM', 
               right_on='Item Number', 
               how='left', 
               suffixes=("", "_temp")).reset_index(drop=True)

cwf['Reference_feature_System_Name(ECCN)'] = 'Export Classification Number (ECCN)'
cwf['Reference feature group ID_(ECCN)'] = cwf['ECCN ']
print(sum(cwf['Reference feature group ID_(ECCN)'].isna()))

cwf['Reference feature group ID_(ECCN)'] = cwf.groupby(['CATALOG_GROUP_ID'], group_keys=False,
                                            sort=False)['Reference feature group ID_(ECCN)'].apply(lambda x: x.fillna(x.mode()[0]))

cwf['FNAME(ECCN)'] = 'Export Classification Number (ECCN)'
cwf['FVALUE(ECCN)'] = cwf['Reference feature group ID_(ECCN)']
cwf['Reference_feature_System_NameCommodityCode'] = 'Commodity Code'
cwf['Reference feature group ID_CommodityCode'] = cwf['Commodity']

cwf['Reference feature group ID_CommodityCode'] = cwf.groupby(['CATALOG_GROUP_ID'], group_keys=False,
                                            sort=False)['Reference feature group ID_CommodityCode'].apply(lambda x: x.fillna(x.mode()[0]))

cwf['FNAME_CommodityCode'] = 'Commodity Code'
cwf['FVALUE_CommodityCode'] = cwf['Reference feature group ID_CommodityCode']

eclass_data = [
    ['ECLASS5_1', 'ECLASS6_0', 'ECLASS6_2', 'ECLASS7_0', 'ECLASS7_1', 'ECLASS10_0_1', 'ECLASS10_1'],
    ['32030000', '21179190', '21179190', '21179190', '21179190', '21179190', '32030000'],
    ['32020000', '27069290', '27069290', '27069290', '27069290', '27069290', '32020100'],
    ['27069290', '27069290', '27069290', '27069290', '27069290', '27069290', '27069290'],
    ['36610408', '36610408', '36610408', '36610408', '36610408', '36610408', '36610408'],
    ['27061003', '27061003', '27061003', '27069290', '27069290', '27069290', '27061003'],
    ['27061803', '27061803', '27061803', '27061803', '27061803', '27061003', '27061803'],
    ['27260801', '27261501', '27261501', '27261501', '27261501', '27261501', '27261790'],
    ['21160190', '21160190', '21160190', '21160190', '21160190', '27261790', '21160190'],
    ['32020103', '32020103', '32020103', '32020103', '32020103', '32020103', '32020103'],
    ['27110636', '27110636', '27110636', '27110636', '27110636', '27110636', '27110636'],
    ['21170590', '32139000', '32139000', '34550000', '34550000', '34550000', '21170590'],
    ['27201304', '27201304', '27201304', '27201304', '27201304', '27201304', '27201304'],
    ['27230218', '27230218', '27230218', '27230218', '27230218', '27230218', '27230218'],
    ['27200307', '27200307', '27200307', '27200307', '27200307', '27200307', '27200307'],
    ['32020100', '32029090', '32029090', '32029090', '32029090', '32029090', '32020100'],
    ['27272704', '27061003', '27061003', '27069290', '27069290', '27069290', '27272704'],
    ['27270905', '27270905', '27270905', '27270905', '27270905', '27270905', '27270905'],
    ['27040000', '19051090', '19051090', '19051090', '19051090', '19051090', '27040000'],
    ['23110690', '23110690', '23110690', '23110690', '23110690', '23110690', '23110690'],
    ['27061802', '27061802', '27061802', '27061802', '27061802', '27061802', '27061802'],
    ['27110635', '27110635', '27110635', '27110635', '27110635', '27110635', '27110635'],
    ['23330201', '23330201', '23330201', '23330201', '23330201', '23330201', '23330201'],
    ['21170503', '21170503', '21170503', '21170503', '21170503', '27273692', '21170503']]


eclass = pd.DataFrame(data=eclass_data, 
                      columns=['ECLASS5_1', 'ECLASS6_0', 'ECLASS6_2', 'ECLASS7_0', 'ECLASS7_1', 'ECLASS10_0_1', 'ECLASS10_1'],
                      dtype=str)

eclass.set_index('ECLASS5_1', inplace=True)

cwf = pd.merge(cwf, eclass, 
               left_on='ECLASS', 
               right_on='ECLASS5_1', 
               how='left', 
               suffixes=("", "_temp")).reset_index(drop=True)

cwf['Reference_feature_System_Name_e5.1'] = 'ECLASS-5.1'
cwf['Reference feature group ID_e5.1'] = cwf['ECLASS']
cwf['FNAMEe5.1'] = 'ECLASS-5.1'
cwf['FVALUEe5.1'] = cwf['ECLASS']

cwf['Reference_feature_System_Namee6.0'] = 'ECLASS-6.0'
cwf['Reference feature group ID_e6.0'] = cwf['ECLASS6_0']
cwf['FNAMEe6.0'] = 'ECLASS-6.0'
cwf['FVALUEe6.0'] = cwf['ECLASS6_0']

cwf['Reference_feature_System_Namee6.2'] = 'ECLASS-6.2'
cwf['Reference feature group ID_e6.2'] = cwf['ECLASS6_2']
cwf['FNAMEe6.2'] = 'ECLASS-6.2'
cwf['FVALUEe6.2'] = cwf['ECLASS6_2']

cwf['Reference_feature_System_Namee7.0'] = 'ECLASS-7.0'
cwf['Reference feature group ID_e7.0'] = cwf['ECLASS7_0']
cwf['FNAMEe7.0'] = 'ECLASS-7.0'
cwf['FVALUEe7.0'] = cwf['ECLASS7_0']

cwf['Reference_feature_System_Namee7.1'] = 'ECLASS-7.1'
cwf['Reference feature group ID_e7.1'] = cwf['ECLASS7_1']
cwf['FNAMEe7.1'] = 'ECLASS-7.1'
cwf['FVALUEe7.1'] = cwf['ECLASS7_1']

cwf['Reference_feature_System_Namee10.0.1'] = 'ECLASS-10.0.1'
cwf['Reference feature group ID_e10.0.1'] = cwf['ECLASS10_0_1']
cwf['FNAMEe10.0.1'] = 'ECLASS-10.0.1'
cwf['FVALUEe10.0.1'] = cwf['ECLASS10_0_1']

cwf['Reference_feature_System_Namee10.1'] = 'ECLASS-10.1'
cwf['Reference feature group ID_e10.1'] = cwf['ECLASS10_1']
cwf['FNAMEe10.1'] = 'ECLASS-10.1'
cwf['FVALUEe10.1'] = cwf['ECLASS10_1']

print('\nSAVING MASTER DATA')
name = "Current Working File_Price {year}{curr}-{month}"
file_suffix = name.format(year=getyear, curr=curr, month=getmonth)

cwf['SUPPLIER_AID'] = cwf['PARTNUM'].reset_index(drop=True)
cwf['MANUFACTURER_AID'] = cwf['PARTNUM']
cwf['PREIS']=cwf['VOLUMEPRICE'].round(2).reset_index(drop=True) #round the volume price

cols = ['UNSPSCCODE', 'SUPPLIER_AID', 'PREIS', 'ORDER_UNIT', 'CONTENT_UNIT', 'Mengeneinheit', 
        'Menge pro Verkaufseinheit/ NO_CU_PER_OU', 'QUANTITY_MIN', 'PRICE_QUANTITY', 'DESCRIPTION_SHORT (80)', 
        'DESCRIPTION_LONG', 'CATALOG_GROUP_ID', 'MANUFACTURER_NAME', 'MANUFACTURER_AID', 'DELIVERY_TIME', 'KEYWORD (max 50)',
        'MIME_SOURCE1', 'MIME_SOURCE1_lrg', 'MIME_SOURCE1_lrg_weblink', 'MIME_PURPOSE1', 'MIME_DESCR1', 'MIME_TYPE1',
        'MIME_SOURCE2', 'MIME_PURPOSE2', 'MIME_TYPE2', 'MIME_DESCR2', 'Reference_feature_System_Name(ECCN)', 
        'Reference feature group ID_(ECCN)', 'FNAME(ECCN)', 'FVALUE(ECCN)', 'Reference_feature_System_NameCommodityCode',
        'Reference feature group ID_CommodityCode', 'FNAME_CommodityCode', 'FVALUE_CommodityCode', 
        'Reference_feature_System_Name_e5.1', 'Reference feature group ID_e5.1', 'FNAMEe5.1', 'FVALUEe5.1',
        'Reference_feature_System_Namee6.0', 'Reference feature group ID_e6.0', 'FNAMEe6.0', 'FVALUEe6.0',
        'Reference_feature_System_Namee6.2', 'Reference feature group ID_e6.2', 'FNAMEe6.2', 'FVALUEe6.2', 
        'Reference_feature_System_Namee7.0', 'Reference feature group ID_e7.0', 'FNAMEe7.0', 'FVALUEe7.0', 
        'Reference_feature_System_Namee7.1', 'Reference feature group ID_e7.1', 'FNAMEe7.1', 'FVALUEe7.1', 
        'Reference_feature_System_Namee10.0.1', 'Reference feature group ID_e10.0.1', 'FNAMEe10.0.1', 'FVALUEe10.0.1',
        'Reference_feature_System_Namee10.1', 'Reference feature group ID_e10.1', 'FNAMEe10.1', 'FVALUEe10.1']

for column in cwf.columns:
    if column not in cols:
        cwf = cwf.drop([column], axis=1)
        print(column, '\tdeleted')

cwf.to_excel(f"{dir_path}\py_{file_suffix}.xlsx", 
             sheet_name='Sheet1', 
             na_rep='', 
             index=False, 
             engine='xlsxwriter')

print('#'*28)

print('\nWRITING LOGS...')
with open("R:\Doc\Kunden Produktlisten\!!!_PY NOTEBOOKS\protocoll.txt", 'a') as f:
    f.write('\n' + str(datetime.today().strftime('%Y-%m-%d')) + '\tCREATED:\t' + file_suffix + '\tIN\t' + dir_path)

print('\nSAMPLING AND PRICE TESTING')
sample = cwf[['SUPPLIER_AID', 'PREIS']].iloc[np.random.randint(low=0, high=int(cwf.shape[0]), size=int(cwf.shape[0]/1000))]

for SUPPLIER_AID in sample['SUPPLIER_AID']:
    webbrowser.open_new('https://www.thorlabs.com/thorproduct.cfm?partnumber={}'.format(SUPPLIER_AID))
    print('\nItem: ', SUPPLIER_AID)
    print('\nPrice: ', sample[sample['SUPPLIER_AID']==SUPPLIER_AID]['PREIS'].item())
    if str(input('\nIs it correct on the page (y/n)? ')) == 'y':
        continue
    else:
        print('Error! correct the file!')
        break

print('\nPROCESS FINISHED.')