# -*- coding: utf-8 -*-
"""
Created on Thu Oct 26 09:03:31 2017

@author: Andy
"""

#%%
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import math

df = pd.read_csv('Laser Module Consumption Monthly.csv')
summary = df.describe()
means = summary.iloc[1, :]
channels = pd.Series(df.columns[1:])

#%%
# Visualizing order history
fig = plt.figure(figsize=(15, 6))
means.plot(kind='bar')
plt.title('Mean Laser Consumption')
plt.xlabel('ITU Channel')
plt.ylabel('Average Monthly Consumpmtion')
plt.savefig('Figures\\Average_Monthly_Consumption_By_Channel.png')

#%%
def monthly_consumption( df ):
  #Split by years with same month number
  years = {
     '2015': None,
     '2016': None,
     '2017': None
          }
  years['2015'] = df.iloc[0:3, :]
  years['2016'] = df.iloc[3:15, :]
  years['2017'] = df.iloc[15:, :]

  years['2016'] = years['2016'].reset_index()
  years['2016'] = years['2016'].drop('index', axis=1)
  years['2017'] = years['2017'].reset_index()
  years['2017'] = years['2017'].drop('index', axis=1)

  for year in years:
      years[ year ]['Month'] = years[ year ]['Month'].map(lambda x: x.split('-')[1]) 


  #Charts for all historical consumption for each channel
  fig2 = plt.figure(figsize=(10, 12))

  for idx, channel in enumerate(channels):
      ax = plt.subplot(8, 6, idx+1)
      plt.tight_layout()
      ax.set_xlim(1, 12)
      plt.title('Channel: {}'.format(channel))
      for year in years:
          plt.plot(years[year]['Month'], years[year][channel])        
          
# Generate plots for rolling average consumptions of lasers of a given type
# with a given window (number of months to be averaged)
def rolling_average( df, window ):
  rolling = pd.rolling_mean(df, window=window)
  channels = df.columns[1:]
  fig = plt.figure(figsize=(15, 12))

  for idx, channel in enumerate(channels):
    ax = plt.subplot(8, 6, idx+1)
    plt.tight_layout()
    rolling[channel].plot(kind='line')
    plt.title('Channel: {}'.format(channel))
  plt.savefig('Figures\\Rolling_Average_Monthly_Laser_Consumption.png')

# Generate plots for histograms of each channel
def channel_histograms( df ):
  channels = df.columns[1:]
  fig = plt.figure(figsize=(15, 12))

  for idx, channel in enumerate(channels):
    ax = plt.subplot(8, 6, idx+1)
    plt.tight_layout()
    plt.hist(df[channel])
    plt.title('Channel: {}'.format(channel))
  plt.savefig('Figures\\Monthly_Laser_Consumption_Histograms.png')
#%%
#Correlate previous month with next month
prevnext = pd.DataFrame(columns=['prev', 'next'])
def apply_prev_next(series):
    for i in series:
        if('val' not in locals()):
            val = i
            continue
        else:
            globals()['prevnext'] = globals()['prevnext'].append({'prev': val, 'next': i}, ignore_index=True)
            val = i

df.iloc[:, 1:].apply(apply_prev_next)

fig = plt.figure(figsize=(15, 6))
plt.scatter(x=prevnext['prev'], y=prevnext['next'])
# prevnext.plot(kind='scatter', x='prev', y='next')
plt.xlabel('Monthly consumption for month n')
plt.ylabel('Monthly consumption for month n+1')
plt.title('Comparison of previous month consumption with current')
plt.savefig('Figures\\Comparison_of_Previous_and_Next_Month_Laser_Consumption.png')
#Largely unhelpful

