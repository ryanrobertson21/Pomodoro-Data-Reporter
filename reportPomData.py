import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import style
from matplotlib.dates import DateFormatter

style.use('fivethirtyeight')

# Read the excel worksheet
allData = pd.read_excel(open('/Users/RyanRobertson21/Desktop/pomTracker.xlsx', 'rb'), sheetname=0)

# Set the index
allData = allData.set_index('DATE')
allData.index = pd.to_datetime(allData.index, unit='s')

# Use this as the index for reports and to get the total poms data
dayAllData = allData.ofHOUR.resample('D', how='sum').round(2)
weekAllData = allData.ofHOUR.resample('W', how='sum').round(2)
monthAllData = allData.ofHOUR.resample('M', how='sum').round(2)
yearAllData = allData.ofHOUR.resample('A', how='sum').round(2)

# Total data from all categories to be tracked
dfTotalPoms = allData.loc[(allData['CATEGORY'] != 'DontTrack')]

# Isolate Programming Data
dfProgramming = allData.loc[(allData['CATEGORY'] == 'Python') | (allData['CATEGORY'] == 'Git') | (allData['CATEGORY'] == 'HTML') | (allData['CATEGORY'] == 'Computer Science') | (allData['CATEGORY'] == 'Linux') | (allData['CATEGORY'] == 'IT')]

# Isolate Other Data (anything that's not programming)
dfOther = allData.loc[(allData['CATEGORY'] != 'Python') & (allData['CATEGORY'] != 'Git') & (allData['CATEGORY'] != 'HTML') & (allData['CATEGORY'] != 'Computer Science') & (allData['CATEGORY'] != 'Linux') & (allData['CATEGORY'] != 'IT') & (allData['CATEGORY'] != 'SQL') & (allData['CATEGORY'] != 'DontTrack')]

# Resample for daily, weekly, monthly and yearly. Also round to 2 decimal places
totalPomDay = dfTotalPoms.ofHOUR.resample('D', how='sum').round(2)
totalPomWeek = dfTotalPoms.ofHOUR.resample('W', how='sum').round(2)
totalPomMonth = dfTotalPoms.ofHOUR.resample('M', how='sum').round(2)
totalPomYear = dfTotalPoms.ofHOUR.resample('A', how='sum').round(2)

progPomDay = dfProgramming.ofHOUR.resample('D', how='sum').round(2)
progPomWeek = dfProgramming.ofHOUR.resample('W', how='sum').round(2)
progPomMonth = dfProgramming.ofHOUR.resample('M', how='sum').round(2)
progPomYear = dfProgramming.ofHOUR.resample('A', how='sum').round(2)

otherPomDay = dfOther.ofHOUR.resample('D', how='sum').round(2)
otherPomWeek = dfOther.ofHOUR.resample('W', how='sum').round(2)
otherPomMonth = dfOther.ofHOUR.resample('M', how='sum').round(2)
otherPomYear = dfOther.ofHOUR.resample('A', how='sum').round(2)


# Print daily data total, programming, and miscellaneous
print('\n        Daily Data\n')
daily = pd.concat([totalPomDay, progPomDay, otherPomDay], axis=1)        # Concatenate the different series of data together
daily.fillna(value=0, inplace=True)                                      # Replace NaNs with 0s
daily.columns = ['Total', ' Prog', ' Misc']                              # Rename the columns
daily.index.rename('', inplace=True)                                     # Get rid of index label
dailyR = daily.loc[::-1]                                                 # Reverse the dataframe and save under new dataframe so I can use old dataframe order for the plots
dailyR.index = dailyR.index.strftime('%-m-%d')                           # Reformat the index to be month (with no leading zero), then day
forDailyPlot = daily.tail(n=14)                                          # Create a non reversed dataframe to plot the last 14 days
print(dailyR.head(n=14))                                                 # Print all daily pom data for the last 14 days

# Print weekly data total, programming, and miscellaneous
print('\n\n       Weekly Data\n')
weekly = pd.concat([totalPomWeek, progPomWeek, otherPomWeek], axis=1)
weekly.fillna(value=0, inplace=True)
weekly.columns = [' Total', '  Prog', ' Misc']
weekly.index.rename('', inplace=True)
weeklyR = weekly.loc[::-1]
weeklyR.index = weeklyR.index.strftime('%-m-%d')
forWeeklyPlot = weekly.tail(n=10)
print(weeklyR.head(n=10))

# Print monthly data total, programming, and miscellaneous
print('\n\n     Monthly Data\n')
monthly = pd.concat([totalPomMonth, progPomMonth, otherPomMonth], axis=1)
monthly.fillna(value=0, inplace=True)
monthly.columns = [' Total', '   Prog', '  Misc']
monthly.index.rename('', inplace=True)
monthlyR = monthly.loc[::-1]
monthlyR.index = monthlyR.index.strftime('%B')
forMonthlyPlot = monthly.tail(n=12)
print(monthlyR.head(n=6))

# Print yearly data total, programming, and miscellaneous
print('\n\n     Yearly Data\n')
yearly = pd.concat([yearAllData, progPomYear, otherPomYear], axis=1)
yearly.fillna(value=0, inplace=True)
yearly.columns = ['Total', 'Prog', 'Misc']
yearly.index.rename('', inplace=True)
yearlyR = yearly.loc[::-1]
yearlyR.index = yearlyR.index.strftime('%Y')
forYearlyPlot = yearly.tail(n=2)
print(yearlyR.head(n=2))

# Print total hours studied programming
allTimeProgHours = progPomYear[0] + progPomYear[1] +progPomYear[2]
print('\n\nTotal time spent studying programming: ', str(allTimeProgHours))

# Print total pom hours for 2016
totalYearPomHours = progPomYear[1] + otherPomYear[1]
print('\n\nTotal Pom hours for 2016: ', str(totalYearPomHours))

# Print total pom hours for 2017
totalYearPomHours = progPomYear[2] + otherPomYear[2]
print('\n\nTotal Pom hours for 2017: ', str(totalYearPomHours))

# Format the dates for plots
DWformatter = DateFormatter('%-m-%d')
Mformatter = DateFormatter('%b')
Yformatter = DateFormatter('%Y')


# Plot the daily data in a stacked bar plot
mpl_fig2 = plt.figure()
ax2 = mpl_fig2.add_subplot(111)
p1 = ax2.bar(forDailyPlot.index, forDailyPlot[' Prog'], width=.75, label = 'Programming')
p2 = ax2.bar(forDailyPlot.index, forDailyPlot[' Misc'],  width=.75, label = 'Other', color=(1.0,0.5,0.62), bottom = forDailyPlot[' Prog'])
plt.gcf().axes[0].xaxis.set_major_formatter(DWformatter)
plt.xticks(forDailyPlot.index, rotation=30)
plt.xlabel('Date')
plt.ylabel('Hours')
plt.title('All Poms 2016 \n Daily')
plt.legend(loc=0)

for tick in ax2.xaxis.get_majorticklabels():
    tick.set_horizontalalignment("left")

plt.show()

# Plot the weekly data in a stacked bar plot
mpl_fig1 = plt.figure()
ax1 = mpl_fig1.add_subplot(111)
p1 = ax1.bar(forWeeklyPlot.index, forWeeklyPlot['  Prog'], width=4, label = 'Programming')
p2 = ax1.bar(forWeeklyPlot.index, forWeeklyPlot[' Misc'],  width=4, label = 'Other', color=(1.0,0.5,0.62), bottom = forWeeklyPlot['  Prog'])
plt.gcf().axes[0].xaxis.set_major_formatter(DWformatter)
plt.xticks(forWeeklyPlot.index, rotation=30)
plt.xlabel('Date')
plt.ylabel('Hours')
plt.title('All Poms 2016 \n Weekly')
plt.legend(loc=0)                                   # Puts the legend in the "best" location

for tick in ax1.xaxis.get_majorticklabels():
      tick.set_horizontalalignment("left")

plt.show()


# Plot the monthly data in a stacked bar plot
mpl_fig2 = plt.figure()
ax2 = mpl_fig2.add_subplot(111)
p1 = ax2.bar(forMonthlyPlot.index, forMonthlyPlot['   Prog'], width=20, label = 'Programming')
p2 = ax2.bar(forMonthlyPlot.index, forMonthlyPlot['  Misc'],  width=20, label = 'Other', color=(1.0,0.5,0.62), bottom = forMonthlyPlot['   Prog'])
plt.gcf().axes[0].xaxis.set_major_formatter(Mformatter)
plt.xticks(forMonthlyPlot.index, rotation=30)
plt.xlabel('Date')
plt.ylabel('Hours')
plt.title('All Poms 2016 \n Monthly')
plt.legend(loc=0)

for tick in ax2.xaxis.get_majorticklabels():
    tick.set_horizontalalignment("left")

plt.show()

# Plot the yearly data in a stacked bar plot
mpl_fig2 = plt.figure()
ax2 = mpl_fig2.add_subplot(111)
p1 = ax2.bar(forYearlyPlot.index, forYearlyPlot['Prog'], width=80, label = 'Programming')
p2 = ax2.bar(forYearlyPlot.index, forYearlyPlot['Misc'],  width=80, label = 'Other', color=(1.0,0.5,0.62), bottom = forYearlyPlot['Prog'])
plt.gcf().axes[0].xaxis.set_major_formatter(Yformatter)
plt.xticks(forYearlyPlot.index, rotation=30)
plt.xlabel('Date')
plt.ylabel('Hours')
plt.title('All Poms \n Yearly')
plt.legend(loc=0)

for tick in ax2.xaxis.get_majorticklabels():
    tick.set_horizontalalignment("left")

plt.show()