#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Nov 14 14:43:56 2022

@author: omar.elkhawass
"""


#%%
import numpy as np
import pandas as pd
import statsmodels.api as sm 
import pandas
import matplotlib.pyplot as plt
import pandas_datareader as pdr


#%%     1) Download Data

NumCompanies = 25

Student_List = pd.read_excel('/Users/omar.elkhawass/Desktop/Student_Tickers.xlsx')
companies = Student_List.loc[Student_List['Full Name'] == 'Elkhawass,Omar Wael', 'Company 1':'Company 25']   #creating a df with my 25 companies
companyList = []    #creating an empty list
for i in range(NumCompanies):
    companyList.append(companies.iloc[:,i].item())    #for loop to index into each column in companies df, convert the object to string, and append into companyList
    
startDate = '2006-01-31'
endDate = '2021-12-31'

companyData = pdr.DataReader(companyList,'yahoo', startDate, endDate)
SP500 = pdr.DataReader('^GSPC', 'yahoo', startDate, endDate)

#DF's for Close/AdjClose/Volume
Price_Daily = companyData['Close']
AdjClose_Daily = companyData['Adj Close']
Volume_Daily = companyData['Volume']
#Export to excel 
writer = pd.ExcelWriter('Omar_Elkhawass_400329748.xlsx', engine = 'xlsxwriter')
Price_Daily.to_excel(writer, sheet_name = 'Price_Daily')
AdjClose_Daily.to_excel(writer, sheet_name = 'AdjClose_Daily')
Volume_Daily.to_excel(writer, sheet_name = 'Volume_Daily')
SP500.to_excel(writer, sheet_name = 'S&P 500')



#%%     2) Calculate Firm Information

# Calculate the size (market capitalization) of each firm per year, export to excel 
SP500_Constituent = pd.read_excel('/Users/omar.elkhawass/Desktop/S&P500_Constituents.xlsx') 
Price_Yearly = (companyData['Close'].resample('1y').ffill())

shares_out = [] 
for i in range(NumCompanies):
    shares_outstanding = SP500_Constituent.loc[SP500_Constituent['ticker'] == companyList[i], 'Share_outstanding'].item()  #finding tickers from SP500 Constituent in company list and getting its shares outstanding value 
    shares_out.append(shares_outstanding)   #adding corresponding shares outstanding numbers into a list    # .item() to convert type from object to float 
    
Market_Cap = Price_Yearly.loc[:, companyList] * np.array(shares_out)   #new df with yearly close prices of each co. multiplied by respective shares outstanding value 

Market_Cap.to_excel(writer, sheet_name = 'Size')


# Sum of daily Volume for each firm per year, export to excel 
Volume_Annual = Volume_Daily.resample('1y').sum()
Volume_Annual.to_excel(writer, sheet_name = 'Volume_Annual')


# Monthly returns, export to excel 
Returns_Monthly = AdjClose_Daily.resample('1m').ffill().pct_change()
Returns_Monthly.to_excel(writer, sheet_name = 'Returns_Monthly')



#%%     3) Summary Statistics of Portfolio

# For each firm, report its Minimum, Maximum, Mean, and volatility of Annual returns
Returns_Yearly = Returns_Monthly.resample('1y').sum()
Returns_Yearly_Stats = Returns_Yearly.describe()
Returns_Yearly_Stats = Returns_Yearly_Stats.drop(['count','25%','50%','75%'])  #drop useless data


# For each firm, report the size (market cap OR market value of equity) at the end of your sample
Returns_Yearly_Stats.loc['Market_Cap'] = Market_Cap.iloc[15]


# Report industry of each company
IndustryList = []    
for i in range(NumCompanies):
    industry =  SP500_Constituent.loc[SP500_Constituent['ticker'] == companyList[i], 'Industry'].item()  #finding tickers from SP500 Constituent in company list and getting its industry 
    IndustryList.append(industry)      #inserting values from previous line into industry list
Returns_Yearly_Stats.loc['Industry'] = IndustryList


# Report Market beta for each firm using the last 5 years (2017:2021)
SP500_Returns = SP500['Adj Close'].resample('1y').ffill().pct_change()    
recent_SP500_Returns = SP500_Returns.tail(5)
recent_Returns_Yearly = Returns_Yearly.tail(5)    #last 5 years only   
    
Returns_Yearly_Stats.loc['Market_Beta'] = 0    
for i in range(NumCompanies):           #using for loop to calculate Beta with Least Squares and fill Market_Beta in stats df
    x = recent_SP500_Returns
    y = recent_Returns_Yearly[companyList[i]]
    regression_model = sm.OLS(y, sm.add_constant(x))
    regression_result = regression_model.fit()
    beta = regression_result.params[1]
    Returns_Yearly_Stats.loc['Market_Beta'].iloc[i] = beta
    
# Compare with the Beta information in sheet S&P 500 Constituents, do they differ by more than 10%?
todaybeta = []
for i in range(NumCompanies):
    SP_beta = SP500_Constituent.loc[SP500_Constituent['ticker'] == companyList[i], 'Beta'].item()  
    todaybeta.append(SP_beta)
Returns_Yearly_Stats.loc['Todays_Beta'] = todaybeta

Returns_Yearly_Stats.loc['Beta_Diff(%)'] = ((Returns_Yearly_Stats.loc['Market_Beta'] - Returns_Yearly_Stats.loc['Todays_Beta']) / Returns_Yearly_Stats.loc['Todays_Beta']) * 100

greater10 = ((Returns_Yearly_Stats.loc['Beta_Diff(%)'] > 10) | (Returns_Yearly_Stats.loc['Beta_Diff(%)'] < -10))
Returns_Yearly_Stats.loc['+ 10% Diff'] = greater10


Returns_Yearly_Stats.to_excel(writer, sheet_name = 'Annual Returns Stats')



#%%     4) Portfolio Analysis

# Portfolio (Method: Equal)
SP500_Monthly_Returns = SP500['Adj Close'].resample('1m').ffill().pct_change().to_frame()
SP500_Monthly_Returns = SP500_Monthly_Returns.iloc[1:, :]   #getting rid of first row which is null

Portfolio = SP500_Monthly_Returns
Portfolio = Portfolio.rename(columns = {'Adj Close':'SP500'})

count_firms = Returns_Monthly.count(axis=1)
weight = 1 / count_firms

Weighted_Returns_Monthly = Returns_Monthly.mul(weight, axis=0)    #multiplying monthly returns by weight 

Portfolio_Returns = Weighted_Returns_Monthly.sum(axis=1)  #summing each months return of each company 
Portfolio['Omar Fund'] = Portfolio_Returns

Portfolio.to_excel(writer, sheet_name = 'PortfolioReturn_monthly')


# Fund Summary (annual)
Fund_summary = Portfolio.resample('1y').sum().describe()     #resampling and summing portfolio returns, then getting stats on it
Fund_summary = Fund_summary.drop(['count','25%','50%','75%'])   #dropping useless info

Fund_summary.loc['Alpha'] = 0
Fund_summary.loc['Beta'] = 0
Fund_summary.loc['R2'] = 0
Fund_summary.loc['SharpeRatio'] = 0
Fund_summary.loc['TreynorRatio'] = 0   #adding empty rows to fill with for loop 
for i in range(2):
    x = SP500_Monthly_Returns.resample('1y').sum()    #resampling to annual 
    y = Portfolio.iloc[:,i].resample('1y').sum()      #resampling to annual
    regression_model = sm.OLS(y, sm.add_constant(x))
    regression_result = regression_model.fit()
        
    alpha = regression_result.params[0]
    Beta = regression_result.params[1]
    R2 = regression_result.rsquared 
    SharpeRatio = y.mean() / y.std()
    TreynorRatio = y.mean() / Beta
    
    Fund_summary.loc['Alpha'].iloc[i] = alpha           #filling rows with the calculated values
    Fund_summary.loc['Beta'].iloc[i] = Beta 
    Fund_summary.loc['R2'].iloc[i] = R2
    Fund_summary.loc['SharpeRatio'].iloc[i] = SharpeRatio
    Fund_summary.loc['TreynorRatio'].iloc[i] = TreynorRatio

Fund_summary.to_excel(writer, sheet_name = 'Fund_summary')


# Industry Compositions
Industry_Compositions = pd.DataFrame(companyList, columns = ['Symbols'])
Industry_Compositions['Percentage(%)'] = weight[endDate] * 100 
Industry_Compositions['Industry'] = IndustryList
Industry_Compositions.to_excel(writer, sheet_name = 'Funds_Holdings_Composition')

Industry_Totals = Industry_Compositions[['Industry','Percentage(%)']].groupby('Industry').sum()
plt.pie(Industry_Totals['Percentage(%)'], labels = Industry_Totals.index, autopct='%1.1f%%', shadow = 'True', textprops={'fontsize': 9})
plt.title('Industry Compositions')
plt.savefig('Ind_Compositionz.png', bbox_inches='tight')    #bbox_inches to save entire figure instead of cutting it   
plt.show()
#plot to excel 
workbook  = writer.book
worksheet = writer.sheets['Funds_Holdings_Composition']
worksheet.insert_image('K3', 'Ind_Compositionz.png')         #inserting image into excel sheet


# Cumulative Portfolio Returns Plot 
initial_fund = 1
cumulativeReturns = Portfolio.copy()
cumulativeReturns['SP500_cum'] = initial_fund * (Portfolio['SP500'] + initial_fund).cumprod()
cumulativeReturns['OmarFund_cum'] = initial_fund * (Portfolio['Omar Fund'] + initial_fund).cumprod()  #calculating the cumulative gains with initial fund using portfolio returns 

plt.plot(cumulativeReturns['SP500_cum'])
plt.plot(cumulativeReturns['OmarFund_cum'])
plt.title('Portfolio Cumulative Returns')
plt.legend(labels = ['SP500','Omar'],frameon = True)
plt.grid()
plt.xlabel('Date')
plt.savefig('Cumulative_Returns.png')
plt.show()

workbook  = writer.book
worksheet = writer.sheets['PortfolioReturn_monthly']
worksheet.insert_image('G3', 'Cumulative_Returns.png')


# Annual Portfolio Returns Plot 
Annual_Portfolio_Returns = Portfolio
Annual_Portfolio_Returns['Year'] = Annual_Portfolio_Returns.index.year
APR_yr = Annual_Portfolio_Returns.groupby('Year').sum()

plt.bar(APR_yr.index - 0.15, APR_yr['SP500'], width = 0.3)
plt.bar(APR_yr.index + 0.15, APR_yr['Omar Fund'], width = 0.3)
plt.xticks(APR_yr.index, rotation = 70)
plt.legend(loc = 'best', labels = ['SP500','Omar'],frameon = True)
plt.grid()
plt.title('Annual Portfolio Returns')
plt.savefig('APR_yr.png')
plt.show()

workbook  = writer.book
worksheet = writer.sheets['PortfolioReturn_monthly']
worksheet.insert_image('G24','APR_yr.png')


# Fund Returns Histogram 
plt.title('Histogram for Portfolio Returns')
plt.hist(Portfolio['SP500'], color='b', bins=20, histtype='step')
plt.hist(Portfolio['Omar Fund'], color='r', bins=20, histtype='step')
plt.legend(labels = ['SP500','Omar'], frameon = True)
plt.ylabel('Frequency')
plt.savefig('Returns_Hist.png')
plt.show()

workbook  = writer.book
worksheet = writer.sheets['Fund_summary']
worksheet.insert_image('F7','Returns_Hist.png')

writer.save()   #saving all excel 'writer' edits




