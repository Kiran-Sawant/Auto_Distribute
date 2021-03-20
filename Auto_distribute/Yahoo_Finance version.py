"""For Symbols Checkout https://finance.yahoo.com/
   Symbols are indicated in parenthesis (...)
   eg. S&P Futures: (ES=F), Gold Futures: (GC=F), etc."""

import pandas as pd
import pandas_datareader as pdr
import datetime as dt
import matplotlib.pyplot as plt
import xlwings as xl

tDeltas = {'1Y': dt.timedelta(days=365),
           '2Y': dt.timedelta(days=365*2),
           '3Y': dt.timedelta(days=365*3),
           '4Y': dt.timedelta(days=365*4),
           '5Y': dt.timedelta(weeks=260.714),
           '10Y': dt.timedelta(weeks=521.429)}

# Getting Symbol, timespan, etc.
name = input("Enter symbol exactly as in Yahoo finance: ")
print("Time Spans: 1Y, 2Y, 3Y, 4Y, 5Y, 10Y")
delta = input("Enter the time span: ").upper()
interval = float(input('Enter interval for O2O bin: '))
interval_2 = float(input('Enter interval for H2L bin: '))

#_________Setting Dates_________#
now = dt.date.today()
start = now - tDeltas[delta]

# reading & formatting CSV
price_data = pdr.DataReader(name, data_source='yahoo', start=start, end=now)
price_data.drop(columns=['Volume', 'Adj Close'], inplace=True)

try:            # If price columns are in string format
    price_data['Open'] = price_data['Open'].str.split(',').str.join('').astype('float')
    price_data['High'] = price_data['High'].str.split(',').str.join('').astype('float')
    price_data['Low'] = price_data['Low'].str.split(',').str.join('').astype('float')
except AttributeError:
    pass
# sorting df in ascending date-order
price_data.sort_index(inplace=True)

# creating open 2open & high 2 low returns column
price_data['O2O %'] = (price_data['Open'].pct_change() * 100).round(3)
price_data['H2L %'] = (((price_data['High'] - price_data['Low'])/price_data['Low']) * 100).round(3)

#____________Probability Distribution______________#
# Creating bin for open 2 open
min = int(price_data['O2O %'].min() - 1)
max = int(price_data['O2O %'].max() + 1)
# Divisor
k = int(abs((min - max)/interval))

# creating bin series
bin = [min]
for i in range(k):
    min += interval
    bin.append(min)

bin_Series = pd.Series(bin).round(3)              # Bin Series

#___________Calculating O2O Probability distribution_____________#
def frequency(array):
    start = array[0]
    end = array[1]
    
    k = price_data[price_data['O2O %'].between(start, end)]
    return k['O2O %'].count()

frequency_series = bin_Series.rolling(2).apply(frequency, raw=True)

frequency_table = pd.concat([bin_Series, frequency_series], axis=1)         # Frequency Table
frequency_table.columns = ['bin', 'Frequency']

# Probability & Cu-probability
prob_Series = (frequency_table['Frequency']/len(price_data['O2O %']) * 100).round(2)
cumprob = ((frequency_table['Frequency']/len(price_data['O2O %']) * 100).round(2)).cumsum()

prob_table = pd.concat([prob_Series, cumprob], axis=1)               # Probability table
prob_table.columns = ['Probability %', 'Cum Probability %']

probability_distribution = pd.concat([frequency_table, prob_table], axis=1)      # probability distribution

#_______________Calculating H2L Probability distribution___________________#
h2lmin = int(price_data['H2L %'].min())
h2lmax = int(price_data['H2L %'].max() + 1)

# Divisor
k2 = int(abs((h2lmax - h2lmin)/interval_2))

# creating bin series
bin2 = [h2lmin]
for i in range(k2):
    h2lmin += interval_2
    bin2.append(h2lmin)

h2l_bin_Series = pd.Series(bin2).round(3)


def frequency2(array):
    start = array[0]
    end = array[1]
    
    k = price_data[price_data['H2L %'].between(start, end)]
    return k['H2L %'].count()

h2l_frequency_series = h2l_bin_Series.rolling(2).apply(frequency2, raw=True)

h2l_frequency_table = pd.concat([h2l_bin_Series, h2l_frequency_series], axis=1)         # Frequency Table
h2l_frequency_table.columns = ['bin', 'Frequency']
h2l_frequency_table.fillna(value=0, inplace=True)

# Probability & Cu-probability
h2l_prob_Series = (h2l_frequency_table['Frequency']/len(price_data['H2L %']) * 100).round(2)
h2l_cumprob = ((h2l_frequency_table['Frequency']/len(price_data['H2L %']) * 100).round(2)).cumsum()

h2l_prob_table = pd.concat([h2l_prob_Series, h2l_cumprob], axis=1)               # Probability table
h2l_prob_table.columns = ['Probability %', 'Cum Probability %']

h2l_probability_distribution = pd.concat([h2l_frequency_table, h2l_prob_table], axis=1)      # probability distribution

#___________Plotting Histograms______________#
fig, axis = plt.subplots(nrows=2, ncols=1, figsize=(6, 8))
price_data['H2L %'].hist(bins=bin2, ax=axis[1])
price_data['O2O %'].hist(bins=bin , ax=axis[0])
fig.subplots_adjust(hspace=0.3)

axis[0].set_title('Open to Open Returns Distribution', fontsize=12)
axis[0].set_xlabel('Returns', fontsize=10)
axis[0].set_ylabel('Frequency', fontsize=10)
axis[1].set_title('High to Low Returns Distribution', fontsize=12)
axis[1].set_xlabel('Returns', fontsize=10)
axis[1].set_ylabel('Frequency', fontsize=10)

#_____________Description______________#

def descriptor(df, col_name):
    
    description = df[col_name].describe().to_dict()
    description['range'] = description['max'] - description['min']
    description['kertosis'] = df[col_name].kurtosis()
    description['skew'] = df[col_name].skew()
    description['median'] = df[col_name].median()
    description['Sample Variance'] = df[col_name].var()
    description.pop('25%')
    description.pop('50%')
    description.pop('75%')
    
    return pd.Series(description)

open_description = descriptor(price_data, 'O2O %')                   # Descriptions
h2l_description = descriptor(price_data, 'H2L %')

#______________Averages________________#
# average open 2 open return
avgO2O = price_data['O2O %'].mean()
avgh2l = price_data['H2L %'].mean()

# boolean Series
positive_mask = price_data['O2O %'] > 0
negative_mask = price_data['O2O %'] < 0

pos_ret = price_data[positive_mask]['O2O %'].mean()
neg_ret = price_data[negative_mask]['O2O %'].mean()
pos_count = price_data[positive_mask]['O2O %'].count()
neg_count = price_data[negative_mask]['O2O %'].count()

avg_df = pd.DataFrame([(pos_ret, pos_count), (neg_ret, neg_count)], columns=['Return', 'Count'])
avg_df.index = ['Pos', 'Neg']
avg_df['Count%'] = avg_df['Count'] / len(price_data)
avg_df['Adjusted ret'] = avg_df['Return'] * avg_df['Count%']          # Averages table

#____________Standard-Deviation______________#
upper = dict()
lower = dict()
for i in range(1, 4):
    upper[i] = (i * price_data['O2O %'].std()) + avgO2O
    lower[i] = avgO2O - (i * price_data['O2O %'].std())

up_std = pd.Series(upper)
low_std = pd.Series(lower)

std_dev = pd.concat([up_std, low_std], axis=1)
std_dev.columns = ['upper', 'lower']

# frequency of Returns under std-dev 1, 2 & 3.
std_1 = len(price_data[price_data['O2O %'].between(std_dev.iloc[0, 1], std_dev.iloc[0, 0])])
std_2 = len(price_data[price_data['O2O %'].between(std_dev.iloc[1, 1], std_dev.iloc[1, 0])])
std_3 = len(price_data[price_data['O2O %'].between(std_dev.iloc[2, 1], std_dev.iloc[2, 0])])
# frequency of Returns as a percentage of sample-size
tup_1 = (std_1, std_1 / len(price_data))
tup_2 = (std_2, std_2 / len(price_data))
tup_3 = (std_3, std_3 / len(price_data))

df = pd.DataFrame([tup_1, tup_2, tup_3], columns=['count', 'percentage'])
df.index = [1, 2, 3]                                           # std-dev table
df.index.name = 'Std_dev'
stdDev_table = pd.concat([std_dev, df], axis=1)

# print(price_data)
# print(probability_distribution)
# print(open_description)
# print(h2l_description)
# print(avg_df)
# print(stdDev_table)

work_book = xl.Book()
wbSheet = work_book.sheets.add(name.strip('.csv'))

wbSheet.range('A9').value = price_data
wbSheet.range('I9').options(index=False).value = probability_distribution
wbSheet.range(f'I{len(bin_Series) + 12}').options(index=False).value = h2l_probability_distribution
wbSheet.range('N10').value = open_description
wbSheet.range(f'N{len(bin_Series) + 13}').value = h2l_description
wbSheet.range('B1').value = "Open 2 Open average returns"
wbSheet.range('C1').value = avgO2O
wbSheet.range('B4').value = "High 2 Low returns"
wbSheet.range('C4').value = avgh2l
wbSheet.range('E1').value = avg_df
wbSheet.range('M1').value = stdDev_table
wbSheet.pictures.add(fig)

work_book.save(name + f" distribution D1 {delta}.xlsx")
