import tkinter as tk
from tkinter import filedialog
import datetime as dt
import pandas as pd
import pandas_datareader as pdr
import matplotlib.pyplot as plt
import xlwings as xl

#__________Global Variables____________#
save_location = None                        # location of distribution .xlsx file
o2o_precision = 0.1                         # intervals in Open to open bin
h2l_precision = 0.1                         # intervals in high to low bin
delta = '1Y'
tDeltas = {'1Y': dt.timedelta(days=365),
           '2Y': dt.timedelta(days=365*2),
           '3Y': dt.timedelta(days=365*3),
           '4Y': dt.timedelta(days=365*4),
           '5Y': dt.timedelta(weeks=260.714),
           '10Y': dt.timedelta(weeks=521.429)}

#___________GUI functions_______________#
def o2o_precision_level(value):
    """Sets the Open to Open bin precision level from Scale widget"""

    global o2o_precision
    o2o_precision = float(value)

def h2l_precision_level(value):
    """Sets the High to Low bin precision level from Scale widget"""

    global h2l_precision
    h2l_precision = float(value)

#________________Distribution Functions_________________#
def frequency(array, dataframe=None, col_name=None):
    """Returns the frequency of values within
    a start & end range, in a given DataFrames column"""

    start = array[0]
    end = array[1]
    
    k = dataframe[dataframe[col_name].between(start, end)]
    return k[col_name].count()

def descriptor(df, col_name):
    """Returns a pd.Series of statistical description of a column of a dataframe."""

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

def distribute():
    """Asks for a save location and file name, and returns an excel file(.xlsx)
    with probability distribution of open to open and high to low prices."""

    global save_location
    global o2o_precision
    global h2l_precision
    global tDeltas
    security_name = symbol_var.get().upper()            # getting the symbol entered by the user.

    # Setting Dates_________#
    now = dt.date.today()
    start = now - tDeltas[delta_var.get()]

    # price_data = pd.read_csv(file_location, usecols=['Date', 'Open', 'High', 'Low'], parse_dates=['Date'], index_col=['Date'])
    price_data = pdr.DataReader(security_name, data_source='yahoo', start=start, end=now)
    price_data.drop(columns=['Volume', 'Adj Close'], inplace=True)

    

    try:            # If price columns are in string format
        price_data['Open'] = price_data['Open'].str.split(',').str.join('').astype('float')
        price_data['High'] = price_data['High'].str.split(',').str.join('').astype('float')
        price_data['Low'] = price_data['Low'].str.split(',').str.join('').astype('float')
    except AttributeError:
        pass

    # sorting df in ascending date-order required to calculate o2o% properly.
    price_data.sort_index(inplace=True)
    
    # creating open 2open & high 2 low returns column
    price_data['O2O %'] = (price_data['Open'].pct_change() * 100).round(3)
    price_data['H2L %'] = (((price_data['High'] - price_data['Low'])/price_data['Low']) * 100).round(3)

    # resorting it ot make it look good
    price_data.sort_index(ascending=False, inplace=True)

    #____________O2O Probability Distribution______________#
    # Creating bin for open 2 open___________#
    # maximum & minimum values
    min_ = price_data['O2O %'].min()
    max_ = price_data['O2O %'].max()

    # 5th largest & 5th smallest value
    largest_5th = price_data['O2O %'].nlargest(5).iloc[-1]
    smallest_5th = price_data['O2O %'].nsmallest(5).iloc[-1]

    # Divisor (number of values in bin)
    k = int(abs((smallest_5th - largest_5th)/o2o_precision))

    # creating bin series
    bin_ = [min_ - o2o_precision, smallest_5th, ]
    for i in range(k + 1):
        smallest_5th += o2o_precision
        bin_.append(smallest_5th)
    bin_.append(max_ + o2o_precision)

    bin_Series = pd.Series(bin_).round(1)              # Bin Series

    frequency_series = bin_Series.rolling(2).apply(frequency, raw=True, kwargs={'dataframe': price_data, 'col_name': 'O2O %'})
    frequency_series.fillna(0, inplace=True)
    frequency_series = frequency_series.astype('int32', copy=False)

    frequency_table = pd.concat([bin_Series, frequency_series], axis=1)         # TODO merge Frequency Table
    frequency_table.columns = ['bin', 'Frequency']

    # Probability & Cu-probability
    prob_Series = (frequency_table['Frequency']/len(price_data['O2O %']) * 100).round(2)
    cumprob = ((frequency_table['Frequency']/len(price_data['O2O %']) * 100).round(2)).cumsum()

    prob_table = pd.concat([prob_Series, cumprob], axis=1)               # Probability table
    prob_table.columns = ['Probability %', 'Cum Probability %']

    probability_distribution = pd.concat([frequency_table, prob_table], axis=1)      # TODO merge probability distribution

    #_______________Calculating H2L Probability distribution___________________#
    # lowest & highest low to high returns
    h2lmin = 0
    h2lmax = price_data['H2L %'].max()

    # 5th largest h2l returns
    h2l_5th_largest = price_data['H2L %'].nlargest(5).iloc[-1]

    # Divisor
    k2 = int(abs(h2l_5th_largest/h2l_precision))

    # creating bin series
    bin2 = [0, ]
    for i in range(k2 + 1):
        h2lmin += h2l_precision
        bin2.append(h2lmin)
    bin2.append(h2lmax + h2l_precision)
    # creating a Series out of bin2 list
    h2l_bin_Series = pd.Series(bin2).round(1)

    # Counting frequency of bin values & making a series of it
    h2l_frequency_series = h2l_bin_Series.rolling(2).apply(frequency, raw=True, kwargs={'dataframe': price_data, 'col_name': 'H2L %'})
    h2l_frequency_series.fillna(0, inplace=True)
    h2l_frequency_series = h2l_frequency_series.astype('int32', copy=False)
    # Creating Bin Frequency table.
    h2l_frequency_table = pd.concat([h2l_bin_Series, h2l_frequency_series], axis=1)         # TODO merge Frequency Table
    h2l_frequency_table.columns = ['bin', 'Frequency']
    h2l_frequency_table.fillna(value=0, inplace=True)

    # Probability & Cu-probability
    h2l_prob_Series = (h2l_frequency_table['Frequency']/len(price_data['H2L %']) * 100).round(2)
    h2l_cumprob = ((h2l_frequency_table['Frequency']/len(price_data['H2L %']) * 100).round(2)).cumsum()

    h2l_prob_table = pd.concat([h2l_prob_Series, h2l_cumprob], axis=1)               # Probability table
    h2l_prob_table.columns = ['Probability %', 'Cum Probability %']

    h2l_probability_distribution = pd.concat([h2l_frequency_table, h2l_prob_table], axis=1)      # TODO merge probability distribution

    #___________Plotting Histograms______________#
    fig, axis = plt.subplots(nrows=2, ncols=1, figsize=(6, 8))
    price_data['H2L %'].hist(bins=bin2, ax=axis[1])
    price_data['O2O %'].hist(bins=bin_ , ax=axis[0])
    fig.subplots_adjust(hspace=0.3)

    axis[0].set_title('Open to Open Returns Distribution', fontsize=12)
    axis[0].set_xlabel('Returns', fontsize=10)
    axis[0].set_ylabel('Frequency', fontsize=10)
    axis[1].set_title('High to Low Returns Distribution', fontsize=12)
    axis[1].set_xlabel('Returns', fontsize=10)
    axis[1].set_ylabel('Frequency', fontsize=10)

    open_description = descriptor(price_data, 'O2O %')                   # TODO merge Descriptions
    h2l_description = descriptor(price_data, 'H2L %')

    #______________Averages________________#
    # average open 2 open return
    avgO2O = price_data['O2O %'].mean().__round__(3)
    avgh2l = price_data['H2L %'].mean().__round__(3)

    # boolean Series
    positive_mask = price_data['O2O %'] > 0
    negative_mask = price_data['O2O %'] < 0

    pos_ret = price_data[positive_mask]['O2O %'].mean().__round__(3)
    neg_ret = price_data[negative_mask]['O2O %'].mean().__round__(3)
    pos_count = price_data[positive_mask]['O2O %'].count()
    neg_count = price_data[negative_mask]['O2O %'].count()

    avg_df = pd.DataFrame([(pos_ret, pos_count), (neg_ret, neg_count)], columns=['Return%', 'Count'])
    avg_df.index = ['Pos', 'Neg']
    avg_df['Count%'] = ((avg_df['Count'] / len(price_data)) * 100).__round__(2)
    avg_df['Adj return%'] = ((avg_df['Return%'] * avg_df['Count%']) / 100).__round__(3)         # TODO merge Averages table

    #____________Standard-Deviation______________#
    upper = dict()
    lower = dict()
    for i in range(1, 4):
        upper[i] = ((i * price_data['O2O %'].std()) + avgO2O).__round__(3)
        lower[i] = (avgO2O - (i * price_data['O2O %'].std())).__round__(3)

    # Creating a Series of upper & lower Std dev
    up_std = pd.Series(upper)
    low_std = pd.Series(lower)
    # Creating a DataFrame of upper & lower Std dev
    std_dev = pd.concat([up_std, low_std], axis=1)
    std_dev.columns = ['upper', 'lower']

    # frequency of Returns under std-dev 1, 2 & 3.
    std_1 = len(price_data[price_data['O2O %'].between(std_dev.iloc[0, 1], std_dev.iloc[0, 0])])
    std_2 = len(price_data[price_data['O2O %'].between(std_dev.iloc[1, 1], std_dev.iloc[1, 0])])
    std_3 = len(price_data[price_data['O2O %'].between(std_dev.iloc[2, 1], std_dev.iloc[2, 0])])
    # frequency of Returns as a percentage of sample-size
    tup_1 = (std_1, ((std_1 / len(price_data)) * 100).__round__(3))
    tup_2 = (std_2, ((std_2 / len(price_data)) * 100).__round__(3))
    tup_3 = (std_3, ((std_3 / len(price_data)) * 100).__round__(3))

    df = pd.DataFrame([tup_1, tup_2, tup_3], columns=['count', 'percentage%'])
    df.index = [1, 2, 3]
    df.index.name = 'Std_dev'
    stdDev_table = pd.concat([std_dev, df], axis=1)                 # TODO merge std-dev table

    # asking save location
    save_location = filedialog.asksaveasfile(filetypes=[('Excel File', '*.xlsx')])
    # file_name = file_location.split('/')[-1].strip('.csv')

    # #________________________Writing to spreadsheet__________________________#
    work_book = xl.Book()
    wbSheet = work_book.sheets.add(security_name)

    wbSheet.range('A9').value = price_data
    wbSheet.range('I9').options(index=False).value = probability_distribution
    wbSheet.range(f'I{len(bin_Series) + 12}').options(index=False).value = h2l_probability_distribution
    wbSheet.range('N10').value = open_description
    wbSheet.range(f'N{len(bin_Series) + 13}').value = h2l_description
    wbSheet.range('B1').value = "Open 2 Open returns %"
    wbSheet.range('C1').value = avgO2O
    wbSheet.range('B4').value = "High 2 Low returns %"
    wbSheet.range('C4').value = avgh2l
    wbSheet.range('E1').value = avg_df
    wbSheet.range('M1').value = stdDev_table
    wbSheet.pictures.add(fig)

    work_book.save(save_location.name + '.xlsx')


#_______Initializing GUI window_________#
mainWindow = tk.Tk()
mainWindow.geometry('360x240')
mainWindow.title('Auto Distribute Y-finance')

#_____Creating widgets______#
# Variables___________#
# errorVar = tk.Variable(mainWindow)
symbol_var = tk.StringVar(mainWindow)
delta_var = tk.StringVar(mainWindow, value='1Y')
# Textbox_____________#
symbol_entry = tk.Entry(mainWindow, textvariable=symbol_var)
# Options menu________#
delta_menu = tk.OptionMenu(mainWindow, delta_var, *[i for i in tDeltas])
# Buttons_____________#
save_btn = tk.Button(mainWindow, text='Distribute', command=distribute)
# Labels______________#
tBoxLabel = tk.Label(mainWindow, text='Enter symbol as in Yahoo Finance: ')
o2o_label = tk.Label(mainWindow, text='Select Open to Open bin precision:')
h2l_label = tk.Label(mainWindow, text='Select High to Low bin precision: ')
blank_label = tk.Label(mainWindow, text='  ')
# error_label = tk.Label(mainWindow, text=errorVar)
# Scales______________#
o2o_scale = tk.Scale(mainWindow, from_=0.1, to=2, resolution=0.1, orient=tk.HORIZONTAL)
h2l_scale = tk.Scale(mainWindow, from_=0.1, to=2, resolution=0.1, orient=tk.HORIZONTAL)

o2o_scale.config(command=o2o_precision_level)
h2l_scale.config(command=h2l_precision_level)

#______Placing widgets________#
tBoxLabel.grid(row=0, column=0)

symbol_entry.grid(row=1, column=0, sticky='ew')

delta_menu.grid(row=1, column=1)

o2o_label.grid(row=2, column=0, columnspan=2)
o2o_scale.grid(row=3, column=0, sticky='ew', columnspan=2)

blank_label.grid(row=4, column=0, columnspan=2)

h2l_label.grid(row=5, column=0, columnspan=2)
h2l_scale.grid(row=6, column=0, sticky='ew', columnspan=2)

save_btn.grid(row=7, column=0, padx=50)

# error_label.grid(row=6, column=0, columnspan=2)

mainWindow.mainloop()

# print(type(symbol_var.get()))
# print(h2l_precision)
# print(o2o_precision)