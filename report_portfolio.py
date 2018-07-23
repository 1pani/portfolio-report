import pandas as pd
import numpy as np
import requests
import matplotlib.pyplot as plt
import quandl
quandl.ApiConfig.api_key = '2QW6T_qzvV8UnvutDVVd'

#C:\Users\hp\Downloads\cust_trans.xlsx

#C:\Users\hp\Downloads\mathew.xlsx
cust_path = input("enter the cust-trans file path - ")
cust_path.replace('\\' , '\\\\')
#print(cust_path)

cust_trans = pd.read_excel(cust_path)
#print(cust_trans)
pi = pd.pivot_table(cust_trans , index=(cust_trans['Fund Name'] , cust_trans['Pur Type']) , values = ['Amount']  , aggfunc=[np.sum])

cust_trans['Date'] = cust_trans['Date'].dt.date
cust_trans['Date'] = cust_trans['Date'].astype(str)
#print(cust_trans)
#cust_trans.drop(cust_trans.iloc[6] , axis=0)# accesscode - b70rxoclf7961y8uih87tuhxlhv5bu9h

#import pandas as pd


response = requests.get("https://api.morningstar.com/v2/service/feed/w41tkz1ekvciq914?accesscode=b70rxoclf7961y8uih87tuhxlhv5bu9h")
#https://api.morningstar.com/v2/service/feed/f9h4xj99ribcw8ih?accesscode=b70rxoclf7961y8uih87tuhxlhv5bu9h
#(type(response))
#(response)
print(response.status_code)
#print(response.content)

list_byte = list(response.content)
#list_byte
#type(list_byte)

str_response = (response.content.decode('utf-8').replace("'" , '"'))

#print(str_response)

df = pd.DataFrame(str_response.split('\n'))
#print(df)

foo = lambda x: pd.Series([i for i in (x.split(','))])
rev = df[0].apply(foo)
#print(rev)

amfi=[]
for i in cust_trans['Fund Name']:
    element = i.split()[-1]
    amfi.append(int(element[1:7]))
    
#print(amfi)

rev = rev.drop(rev.index[0])
rev = rev.drop(rev.index[-1])
#print(rev)
#print(rev.isnull().any())

rev = rev.rename(columns={0:'MStarID' , 1:'ISIN' , 2:'FundName' , 3:'AMFIid' , 4:'DailyNAV'})
rev = rev.drop(5 , axis=1)
#print(rev)

rev = rev.replace(' ' , 0)
rev['AMFIid'] = rev['AMFIid'].astype(str).astype(int)
rev['DailyNAV'] = rev['DailyNAV'].astype(str).astype(float)

d = pd.DataFrame()
for i in amfi:
    d = d.append(rev[rev['AMFIid']==i])
#d.groupby('FundName').sum()
#print(d)

i = cust_trans['Amount']
li_neg = [i for i in cust_trans[cust_trans['Pur Type']=='redemption']['Amount']]
li_neg = [-1*i for i in li_neg] 
#print(sum(li_neg))

for i in li_neg:
    print(cust_trans[cust_trans['Amount'] == abs(i)])

cust_trans.loc[cust_trans['Pur Type']=='redemption' , 'Amount'] *= -1
#print(cust_trans)

#path = r'C:\Users\hp\Downloads\PrimaryDataset_MethodologyV120180704.xlsx'
#path.replace('\\' , '\\\\')
#primary = pd.read_excel('C:\Users\hp\Downloads\PrimaryDataset_MethodologyV120180704.xlsx')
path = input("enter your path - ")
path.replace('\\' , '\\\\')
primary = pd.read_excel(path)
print(path)
f = pd.DataFrame()
for i in d['ISIN']:
    f = f.append(primary[primary['ISIN']==i])

#print(d)
d = d.set_index(cust_trans.index)
d['Pur Type'] = cust_trans['Pur Type']
d['Amount'] = cust_trans['Amount']
piv_d = pd.pivot_table(d, index=(d['FundName'] , d['Pur Type']) , values = ['Amount']  , aggfunc=[np.sum])
d = d.groupby(['MStarID' , 'ISIN' , 'FundName' , 'AMFIid' , 'DailyNAV'])['Amount'].sum().reset_index()

pri_d = pd.DataFrame()
for i in d['ISIN']:
    pri_d = pri_d.append(primary[primary['ISIN'] == i])

market_cap = [i for i in pri_d.columns if 'MarketCap' in i]
equity_sector = [i for i in pri_d.columns if 'EquitySector'  in i]
credit = [i for i in pri_d.columns if 'CreditQual'  in i]
asset_alloc = [i for i in pri_d.columns if 'AssetAlloc'  in i]

pri_d = pri_d.replace(' ' , 0)
for i in market_cap:
    pri_d[i] = pri_d[i].astype(str).astype(float)
for i in equity_sector:
    pri_d[i] = pri_d[i].astype(str).astype(float)
for i in credit:
    pri_d[i] = pri_d[i].astype(str).astype(float)
for i in asset_alloc:
    pri_d[i] = pri_d[i].astype(str).astype(float)

d[market_cap[0]] = (pri_d[market_cap[0]]).values
d = d.drop(['MarketCapGiantLongRescaled'] , axis=1)

d_copy = d.copy(deep=True)
for i in range(0,len(market_cap)):
    d_copy[market_cap[i]] = pri_d[market_cap[i]].values
for i in range(0,len(equity_sector)):
    d_copy[equity_sector[i]] = pri_d[equity_sector[i]].values
for i in range(0,len(credit)):
    d_copy[credit[i]] = pri_d[credit[i]].values
for i in range(0,len(asset_alloc)):
    d_copy[asset_alloc[i]] = pri_d[asset_alloc[i]].values

#print(d_copy)
#print(d_copy.columns)
d_copy['MarketCapGiantLongRescaled'] = ((d_copy['Amount'] * d_copy['MarketCapGiantLongRescaled']))/100
#print(d_copy)

d_copy['MarketCapLargeLongRescaled'] = ((d_copy['Amount'] * d_copy['MarketCapLargeLongRescaled']))/100
d_copy['MarketCapMicroLongRescaled'] = ((d_copy['Amount'] * d_copy['MarketCapMicroLongRescaled']))/100
d_copy['MarketCapMidLongRescaled'] = ((d_copy['Amount'] * d_copy['MarketCapMidLongRescaled']))/100
d_copy['MarketCapSmallLongRescaled'] = ((d_copy['Amount'] * d_copy['MarketCapSmallLongRescaled']))/100
#done for market cap

mcap_graph = []
mcap_graph.append((d_copy['MarketCapMidLongRescaled'].sum())*100/(d_copy['Amount']).sum())
mcap_graph.append((d_copy['MarketCapGiantLongRescaled'].sum() + d_copy['MarketCapLargeLongRescaled'].sum())*100/(d_copy['Amount']).sum())
mcap_graph.append((d_copy['MarketCapMicroLongRescaled'].sum() + d_copy['MarketCapSmallLongRescaled'].sum())*100/d_copy['Amount'].sum())
#done for mcap

d_copy['EquitySectorBasicMaterialsLongRescaled'] = ((d_copy['Amount'] * d_copy['EquitySectorBasicMaterialsLongRescaled']))/100
d_copy['EquitySectorCommunicationServicesLongRescaled'] = ((d_copy['Amount'] * d_copy['EquitySectorCommunicationServicesLongRescaled']))/100
d_copy['EquitySectorConsumerCyclicalLongRescaled'] = ((d_copy['Amount'] * d_copy['EquitySectorConsumerCyclicalLongRescaled']))/100
d_copy['EquitySectorConsumerDefensiveLongRescaled'] = ((d_copy['Amount'] * d_copy['EquitySectorConsumerDefensiveLongRescaled']))/100
d_copy['EquitySectorEnergyLongRescaled'] = ((d_copy['Amount'] * d_copy['EquitySectorEnergyLongRescaled']))/100
d_copy['EquitySectorFinancialServicesLongRescaled'] = ((d_copy['Amount'] * d_copy['EquitySectorFinancialServicesLongRescaled']))/100
d_copy['EquitySectorHealthcareLongRescaled'] = ((d_copy['Amount'] * d_copy['EquitySectorHealthcareLongRescaled']))/100
d_copy['EquitySectorIndustrialsLongRescaled'] = ((d_copy['Amount'] * d_copy['EquitySectorIndustrialsLongRescaled']))/100
d_copy['EquitySectorRealEstateLongRescaled'] = ((d_copy['Amount'] * d_copy['EquitySectorRealEstateLongRescaled']))/100
d_copy['EquitySectorTechnologyLongRescaled'] = ((d_copy['Amount'] * d_copy['EquitySectorTechnologyLongRescaled']))/100
d_copy['EquitySectorUtilitiesLongRescaled'] = ((d_copy['Amount'] * d_copy['EquitySectorUtilitiesLongRescaled']))/100
#done for equity sector

equity_graph=[]
equity_graph.append((d_copy['EquitySectorBasicMaterialsLongRescaled'].sum())*100/(d_copy['Amount']).sum())
equity_graph.append((d_copy['EquitySectorCommunicationServicesLongRescaled'].sum())*100/(d_copy['Amount']).sum())
equity_graph.append((d_copy['EquitySectorConsumerCyclicalLongRescaled'].sum())*100/(d_copy['Amount']).sum())
equity_graph.append((d_copy['EquitySectorConsumerDefensiveLongRescaled'].sum())*100/(d_copy['Amount']).sum())
equity_graph.append((d_copy['EquitySectorEnergyLongRescaled'].sum())*100/(d_copy['Amount']).sum())
equity_graph.append((d_copy['EquitySectorFinancialServicesLongRescaled'].sum())*100/(d_copy['Amount']).sum())
equity_graph.append((d_copy['EquitySectorHealthcareLongRescaled'].sum())*100/(d_copy['Amount']).sum())
equity_graph.append((d_copy['EquitySectorIndustrialsLongRescaled'].sum())*100/(d_copy['Amount']).sum())
equity_graph.append((d_copy['EquitySectorRealEstateLongRescaled'].sum())*100/(d_copy['Amount']).sum())
equity_graph.append((d_copy['EquitySectorTechnologyLongRescaled'].sum())*100/(d_copy['Amount']).sum())
equity_graph.append((d_copy['EquitySectorUtilitiesLongRescaled'].sum())*100/(d_copy['Amount']).sum())

d_copy['AssetAllocBondNet'] = ((d_copy['Amount'] * d_copy['AssetAllocBondNet']))/100
d_copy['AssetAllocCashNet'] = ((d_copy['Amount'] * d_copy['AssetAllocCashNet']))/100
d_copy['AssetAllocEquityNet'] = ((d_copy['Amount'] * d_copy['AssetAllocEquityNet']))/100
#done for asset alloc

asset_graph=[]
asset_graph.append((d_copy['AssetAllocBondNet'].sum())*100/(d_copy['Amount']).sum())
asset_graph.append((d_copy['AssetAllocCashNet'].sum())*100/(d_copy['Amount']).sum())
asset_graph.append((d_copy['AssetAllocEquityNet'].sum())*100/(d_copy['Amount']).sum())

d_copy['CreditQualA'] = ((d_copy['Amount'] * d_copy['CreditQualA']))/100
d_copy['CreditQualAA'] = ((d_copy['Amount'] * d_copy['CreditQualAA']))/100
d_copy['CreditQualAAA'] = ((d_copy['Amount'] * d_copy['CreditQualAAA']))/100
#done for credit

credit_graph =[]
credit_graph.append((d_copy['CreditQualA'].sum())*100/(d_copy['Amount']).sum())
credit_graph.append((d_copy['CreditQualAA'].sum())*100/(d_copy['Amount']).sum())
credit_graph.append((d_copy['CreditQualAAA'].sum())*100/(d_copy['Amount']).sum())

colors = ['lightcoral' , 'mediumslateblue' , 'seagreen']

plt.pie(credit_graph , labels=['A' , 'AA' , 'AAA'] , autopct='%.2f' , colors = colors)
plt.title('Credit Exposure', fontsize=20)
plt.savefig('credit.png')
plt.axis('equal')
plt.show()

plt.pie(asset_graph , labels=['BondNet' , 'CashNet' , 'EquityNet'] , autopct='%.2f', colors = colors)
plt.title('Asset Allocation', fontsize=20)
plt.savefig('asset.png')
plt.axis('equal')
plt.show()

plt.pie(equity_graph , labels=['BasicMat' , 'CommSer' , 'ConsCyc' , 'DefLong' , 'EnergyLong' , 'FinancialServ' , 'HealthCare' , 'Indus' , 'RealEs' , 'Tech' , 'Util'],  autopct='%.2f' , colors = colors)
plt.title('Sectoral', fontsize=20)
plt.savefig('sectoral.png')
plt.axis('equal')
plt.show()

plt.pie(mcap_graph , labels=['Mid' , 'Large' , 'Small'] , autopct='%.2f' , colors = colors)
plt.title('Marketcap', fontsize=20)
plt.savefig('marketcap.png')
plt.axis('equal')
plt.show()

#primary[['ISIN' , 'LegalName' , 'SortinoRatio3Yr' ,'SortinoRatio3Yr' , 'SharpeRatio3Yr' , 'TreynorRatio3Yr'  , 'ModifiedDurationLong' , 'YieldToMaturity']]

risk_met = pri_d.copy(deep=True)
risk_met = risk_met.reset_index()
risk_met = risk_met.drop(['index'] , axis=1)
#print(d_copy)
risk_met = risk_met[['MStarID' , 'ISIN' , 'LegalName' , 'AMFICode' , 'SortinoRatio3Yr' , 'SharpeRatio3Yr' , 'TreynorRatio3Yr' , 'ModifiedDurationLong' , 'YieldToMaturity']]
#print(risk_met)
risk_met = risk_met.replace(' ' , 0)
risk_met['AMFICode'] = risk_met['AMFICode'].astype(str).astype(int)
#risk_met['AMFICode'].astype(str).astype(int)
#risk_met.dtypes
risk_met['SortinoRatio3Yr'] = risk_met['SortinoRatio3Yr'].astype(str).astype(float)
risk_met['TreynorRatio3Yr'] = risk_met['TreynorRatio3Yr'].astype(str).astype(float)
risk_met['SharpeRatio3Yr'] = risk_met['SharpeRatio3Yr'].astype(str).astype(float)
risk_met['ModifiedDurationLong'] = risk_met['ModifiedDurationLong'].astype(str).astype(float)
risk_met['YieldToMaturity'] = risk_met['YieldToMaturity'].astype(str).astype(float)

risk_met['Amount'] = d_copy['Amount']
risk_met['DailyNAV'] = d_copy['DailyNAV']
pd.options.mode.chained_assignment = None
#print(risk_met)

hist_return = pri_d.copy(deep=True)
hist_return = hist_return.reset_index()
hist_return = hist_return.drop(['index'] , axis=1)
hist_return = hist_return[['MStarID' , 'ISIN' , 'LegalName' , 'AMFICode' , 'Return1Mth' , 'Return3Mth' , 'Return6Mth' , 'Return1Yr' , 'Return3Yr' , 'Return5Yr']]
hist_return['Amount'] = d_copy['Amount']
hist_return['DailyNAV'] = d_copy['DailyNAV']
#print(hist_return)

fund_graph = d_copy.copy(deep=True)
primary_graph = primary[['MStarID' , 'ISIN' , 'LegalName' , 'MarketCapGiantLongRescaled' , 'MarketCapLargeLongRescaled' , 'MarketCapMidLongRescaled' , 'MarketCapSmallLongRescaled' , 'MarketCapMicroLongRescaled' , 'CreditQualAAA' , 'CreditQualAA' , 'CreditQualA' , 'EquitySectorBasicMaterialsLongRescaled' , 'EquitySectorCommunicationServicesLongRescaled' , 'EquitySectorConsumerCyclicalLongRescaled' , 'EquitySectorConsumerDefensiveLongRescaled' , 'EquitySectorEnergyLongRescaled' , 'EquitySectorFinancialServicesLongRescaled' , 'EquitySectorHealthcareLongRescaled' , 'EquitySectorIndustrialsLongRescaled' , 'EquitySectorRealEstateLongRescaled' , 'EquitySectorTechnologyLongRescaled' , 'EquitySectorUtilitiesLongRescaled' , 'AssetAllocBondNet' , 'AssetAllocCashNet' , 'AssetAllocEquityNet']]
primary_graph = primary_graph.replace(' ' , 0)
for i in primary_graph.columns[3:]:
    primary_graph[i] = primary_graph[i].astype(str).astype(float)

hold_name = ([i for i in primary.columns if 'HoldingDetail_Name' in i])
#fund_hold_name = ([i for i in fund_pri.columns if 'HoldingDetail_Name' in i])
#print('\n')
hold_weight = ([i for i in primary.columns if 'HoldingDetail_Weighting' in i])
#fund_hold_weight = ([i for i in fund_pri.columns if 'HoldingDetail_Weighting' in i])
port_hold = pri_d.copy(deep=True)
port_hold = port_hold.set_index(d_copy.index)

#fund_hold = fund_pri[fund_pri['ISIN'] == isin]
#fund_hold = fund_hold[['MStarID' , 'ISIN' , 'LegalName' , 'AMFICode']]
port_hold=d_copy.copy(deep=True)
for i in port_hold.columns[6:]:
    port_hold = port_hold.drop(i , axis=1)
port_hold = port_hold.set_index(pri_d.index)
for i , j in zip(hold_name , hold_weight):
    port_hold[i] = pri_d[i]
    port_hold[j] = pri_d[j]

#print(port_hold)
port_hold = port_hold.set_index(d_copy.index)
for i in hold_weight:
    port_hold[i] = port_hold[i].replace(' ' , 0)
    port_hold[i] = port_hold[i].astype(str).astype(float)
for i in hold_weight:
    port_hold[i] = (port_hold[i]*port_hold['Amount'])
for i in hold_weight:
    port_hold[i] = (port_hold[i])/100

#print(port_hold)
for i in hold_weight:
    port_hold[i] = ((port_hold[i]*100)/(port_hold['Amount'].sum()))

l_weight = []
l_name = []
for i in range(0,len(port_hold)):
    for j in hold_weight:
        l_weight.append(port_hold[j].iloc[i])
        
for i in range(0,len(port_hold)):
    for j in hold_name:
        l_name.append(port_hold[j].iloc[i])        

s = pd.Series(sorted(zip(l_weight, l_name), reverse=True)[:len(l_name)])
top10 = pd.DataFrame(s.values)
top10 = top10[0].apply(pd.Series)
top10 = top10.rename(columns={0:'Weight%' , 1:'Holding'})
#top10
top10 = top10.groupby('Holding').sum()
top10 = top10.nlargest(10 , 'Weight%')
top10 = top10.round({'Weight%':1})

#print(top10)

#print(cust_trans)

scrape=pd.DataFrame()
for i in cust_trans['Fund Name'].unique():
    scrape = scrape.append(cust_trans[cust_trans['Fund Name']==i].reset_index())
#print(scrape)

for i in scrape['index']:
    print((cust_trans.iloc[i , 0]))
    
amfi_scrape=[]
for i in scrape['Fund Name']:
    element = i.split()[-1]
    amfi_scrape.append(int(element[1:7]))    

cust_trans['AMFI'] = pd.DataFrame(pd.Series(amfi).values)
cust_trans['AMFI'] = cust_trans['AMFI'].astype(str)

#qu = quandl.get('AMFI/'+str(code), start_date='2014-07-09', end_date='2018-07-09')
that_nav=[]
for i in cust_trans[['Date' , 'AMFI']].values:
    print(i[0] , i[1])
    that_nav.append(quandl.get('AMFI/'+str(i[1]), start_date=i[0], end_date=i[0])['Net Asset Value'].values)
    
that_nav_list =[]
#cust_trans[['Date' , 'AMFI']]
for i in range(0,len(that_nav)):
    that_nav_list.append(that_nav[i][0])

particular_nav = pd.Series(that_nav_list)
#print(particular_nav.values)
cust_trans['Purchase NAV'] = particular_nav.values
cust_trans['Units'] = cust_trans['Amount']/cust_trans['Purchase NAV']
#print(cust_trans)

prev = input("enter the previous day in YYYY-MM-DD")
today_nav = []
for i in cust_trans['AMFI']:
    today_nav.append(quandl.get('AMFI/'+str(i), start_date=prev, end_date=prev)['Net Asset Value'].values)
today_nav_list =[]
#today_nav
for i in range(0,len(today_nav)):
    today_nav_list.append(today_nav[i][0])
#print(today_nav_list)

today = pd.Series(today_nav_list)
#print(today.values)
cust_trans['Today NAV'] = today.values

cust_trans['Current Value'] = cust_trans['Today NAV']*cust_trans['Units']
cust_trans.groupby('Fund Name' ).sum().reset_index()
for i in cust_trans['AMFI'].unique():
    print(cust_trans[cust_trans['AMFI']==i]['Amount'].sum())

cust_final = pd.DataFrame()
cust_final = cust_final.append(cust_trans[cust_trans['Pur Type']=='fresh'].reset_index())
cust_final = cust_final.drop('index' , axis=1)
a = []
for i in cust_final['Fund Name'].unique():
    a.append(cust_trans[cust_trans['Fund Name']==i]['Amount'].sum())

cust_final['Amount'] = pd.Series(a).values
#print(cust_final)
a = []
for i in cust_final['Fund Name'].unique():
    a.append(cust_trans[cust_trans['Fund Name']==i]['Current Value'].sum())
cust_final['Current Value'] = pd.Series(a).values
cust_final = cust_final.drop(['Goal Name' , 'Folio' , 'BSE Order No' , 'Status' , 'Pur Type' , 'Order Type' , 'SIP Date' , 'Units' ] , axis=1)
cust_final = cust_final.rename(columns={'Amount':'Purchase Value(INR.)' , 'Purchase NAV':'Purchase NAV(INR.)' , 'Today NAV':'Today NAV(INR.)' , 'Current Value':'Current Value(INR.)'})
cust_final['Absolute Returns(%)'] = ((cust_final['Current Value(INR.)']-cust_final['Purchase Value(INR.)'])*100)/cust_final['Purchase Value(INR.)']
cust_final = cust_final.drop(['AMFI'] , axis=1)
#print(cust_final)
cust_final = cust_final.round({'Purchase NAV(INR.)':2 , 'Today NAV(INR.)':2 , 'Current Value(INR.)':0 , 'Absolute Returns(%)':1})
###################calculation for risk metrics################################


sortino = sum(risk_met['Amount']*risk_met['SortinoRatio3Yr'])/sum(risk_met['Amount'])
sharpe = sum(risk_met['Amount']*risk_met['SharpeRatio3Yr'])/sum(risk_met['Amount'])
treynor = sum(risk_met['Amount']*risk_met['TreynorRatio3Yr'])/sum(risk_met['Amount'])
modified = sum(risk_met['Amount']*risk_met['ModifiedDurationLong'])/sum(risk_met['Amount'])
ytm = sum(risk_met['Amount']*risk_met['YieldToMaturity'])/sum(risk_met['Amount'])

ytm = ("%.2f"%(ytm))
modified = ("%.2f"%(modified))
treynor = ("%.2f"%(treynor))
sharpe = ("%.2f"%(sharpe))
sortino = ("%.2f"%(sortino))

#sortino , sharpe , treynor , modified , ytm
#also do for stddev , avgmaturity , beta

xirr = pd.DataFrame()
for i in cust_trans['Fund Name'].unique():
    xirr = xirr.append(cust_trans[cust_trans['Fund Name']==i].reset_index())
      
xirr = xirr.drop(['index' , 'Goal Name' , 'Folio' , 'BSE Order No' , 'Status'] , axis=1)
xirr = xirr.drop(['Pur Type' , 'Order Type' , 'SIP Date'] , axis=1)
#xirr = xirr.drop(['AMFI' , 'Today NAV'])
xirr['Amount']*=-1
#print(xirr)



writer = pd.ExcelWriter('dupli.xlsx' , engine='xlsxwriter')
workbook = writer.book
worksheet = workbook.add_worksheet('Sheet1')
writer.sheets['Sheet1'] = worksheet
cell_format = workbook.add_format()
cell_format.set_font_color('red')
bold = workbook.add_format()
cell_format.set_bold()
bold.set_bold()
fmt = workbook.add_format()
fmt.set_bg_color('white')
#color_format1 = workbook.add_format({'bg_color' : '#FFC7CE'})
#color_format2 = workbook.add_format({'bg_color' : '#00C7CE'})

for i in range(0,300):
    worksheet.set_row(i,None,fmt)

cust_final.to_excel(writer , sheet_name = 'Sheet1' , startrow = 0 , startcol = 0)

worksheet.write(len(d)+4 ,0 ,  'Portfolio' , cell_format)
worksheet.insert_image(len(d)+5 ,0 ,'marketcap.png', {'x_scale': 0.5, 'y_scale': 0.5})
worksheet.insert_image(len(d)+5 ,5 , 'asset.png', {'x_scale': 0.5, 'y_scale': 0.5})
worksheet.insert_image(len(d)+5 ,10, 'credit.png', {'x_scale': 0.5, 'y_scale': 0.5})
worksheet.insert_image(len(d)+5 ,15 ,  'sectoral.png', {'x_scale': 0.5, 'y_scale': 0.5})
worksheet.write('A1' , 'Performance' , cell_format)

worksheet.write(len(d)+19 ,0, 'Risk metrics' , cell_format)
worksheet.write(len(d)+20+len(d)+4,0, 'Top10 holdings' , cell_format)
#risk_met.to_excel(writer , sheet_name = 'Sheet1' , startrow = len(d)+19 , startcol = 0)
worksheet.write('A24', 'Sortino' , cell_format )
worksheet.write('B24' , sortino)
worksheet.write('A25' , 'Sharpe' , cell_format )
worksheet.write('B25' , sharpe)
worksheet.write('D24' , 'Treynor' , cell_format )
worksheet.write('E24' , treynor)
worksheet.write('D25' , 'Modified' , cell_format )
worksheet.write('E25' , modified)
worksheet.write('G24' , 'YTM' , cell_format )
worksheet.write('H24' , ytm)
worksheet.write('P32' , 'Current date is' + str(prev) , bold)
top10.to_excel(writer , sheet_name = 'Sheet1' , startrow = len(d)+20+len(d)+5 , startcol = 0)
#xirr.to_excel(writer , sheet_name = 'Sheet1' , startrow=31 , startcol=5)

for i in range(0,len(xirr['Fund Name'].unique())):
    
    xirr[xirr['Fund Name']==xirr['Fund Name'].unique()[i]][['Date' , 'Fund Name' , 'Amount']].to_excel(writer , sheet_name = 'Sheet1' , startrow=32 , startcol=5+(5*i))
#for i in range(0,len(d)+5 , 2):
    
#    worksheet.set_row(i, cell_format=color_format1)
#    worksheet.set_row(i + 1, cell_format=color_format2)
worksheet.hide_gridlines()
writer.save()

input("Press enter to finish the program")

