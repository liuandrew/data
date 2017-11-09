# -*- coding: utf-8 -*-
"""
Created on Wed Nov  8 12:18:05 2017

@author: Andy
"""

#%%
from automation1.mssql import DatabaseManagerSQL
from automation1.utilities import *
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os.path

# Connect to SQL
sql = DatabaseManagerSQL()

#%%
# Histogram date between tests, ignoring Batch='1', Chip='1', because that seems to be
# a recurring test unit that leads to lots of negative values 
query1 = "SELECT (floor(DateDiff(d, C.Timestamp, Laser.MaxDate) / 10.00) * 10) as DeltaDate, COUNT(*) as Count \
  FROM OrtelTE.dbo.ChirpSpecTrumData AS C \
    INNER JOIN (  \
      SELECT Batch, Chip, max(Timestamp) as MaxDate  \
      FROM OrtelTe.dbo.ChirpSpecTrumData \
      GROUP BY Batch, Chip  \
    ) AS CMaxDate  \
      ON C.Batch = CMaxDate.Batch AND C.Chip = CMaxDate.Chip AND C.Timestamp = CMaxDate.MaxDate  \
	INNER JOIN (\
		SELECT Device_SN, BatchNo, ChipNo, max(TEST_DT) as MaxDate\
		FROM OrtelTE.dbo.ModuleDistortion\
		GROUP BY Device_SN, BatchNo, ChipNo\
	) AS Laser\
	ON C.Batch = Laser.BatchNo AND C.Chip = Laser.ChipNo\
	WHERE C.Batch != '1' AND C.Chip != '1'\
	GROUP BY (floor(DateDiff(d, C.Timestamp, Laser.MaxDate) / 10.00) * 10)\
	ORDER BY (floor(DateDiff(d, C.Timestamp, Laser.MaxDate) / 10.00) * 10)"

# Histogram date between chip test date and today's date where chip is not found
# in laser sql table, i.e. chip has not been selected
query2 = "SELECT (floor(DateDiff(d, C.Timestamp, CURRENT_TIMESTAMP) / 10.00) * 10) as DeltaDate, COUNT(*) as Count \
  FROM OrtelTE.dbo.ChirpSpecTrumData AS C  \
    INNER JOIN (   \
      SELECT Batch, Chip, max(Timestamp) as MaxDate   \
      FROM OrtelTe.dbo.ChirpSpecTrumData  \
      GROUP BY Batch, Chip   \
    ) AS CMaxDate   \
      ON C.Batch = CMaxDate.Batch AND C.Chip = CMaxDate.Chip AND C.Timestamp = CMaxDate.MaxDate   \
	LEFT JOIN ( \
		SELECT Device_SN, BatchNo, ChipNo, max(TEST_DT) as MaxDate \
		FROM OrtelTE.dbo.ModuleDistortion \
		GROUP BY Device_SN, BatchNo, ChipNo \
	) AS Laser \
	ON C.Batch = Laser.BatchNo AND C.Chip = Laser.ChipNo \
	WHERE C.Batch != '1' AND C.Chip != '1' AND Laser.Device_SN IS NULL \
	GROUP BY (floor(DateDiff(d, C.Timestamp, CURRENT_TIMESTAMP) / 10.00) * 10) \
	ORDER BY (floor(DateDiff(d, C.Timestamp, CURRENT_TIMESTAMP) / 10.00) * 10)"
  
#%%
# Histogram for matched lasers
res1 = sql.ExecQuery(query1)
df1 = pd.DataFrame(data=res1, columns=['DeltaDateBin', 'Count'])
df1 = df1[df1['DeltaDateBin'] > -1]

# Histogram for unused cobs
res2 = sql.ExecQuery(query2)
df2 = pd.DataFrame(data=res2, columns=['DeltaDateBin', 'Count'])

#%%
# Apparently we can just get all the delta dates like whatever
query3 = "SELECT DateDiff(d, C.Timestamp, CURRENT_TIMESTAMP) as DeltaDate \
  FROM OrtelTE.dbo.ChirpSpecTrumData AS C  \
    INNER JOIN (   \
      SELECT Batch, Chip, max(Timestamp) as MaxDate   \
      FROM OrtelTe.dbo.ChirpSpecTrumData  \
      GROUP BY Batch, Chip   \
    ) AS CMaxDate   \
      ON C.Batch = CMaxDate.Batch AND C.Chip = CMaxDate.Chip AND C.Timestamp = CMaxDate.MaxDate   \
	LEFT JOIN ( \
		SELECT Device_SN, BatchNo, ChipNo, max(TEST_DT) as MaxDate \
		FROM OrtelTE.dbo.ModuleDistortion \
		GROUP BY Device_SN, BatchNo, ChipNo \
	) AS Laser \
	ON C.Batch = Laser.BatchNo AND C.Chip = Laser.ChipNo \
	WHERE C.Batch != '1' AND C.Chip != '1' AND Laser.Device_SN IS NULL "
  
query4 = "SELECT DateDiff(d, C.Timestamp, Laser.MaxDate) as DeltaDate \
  FROM OrtelTE.dbo.ChirpSpecTrumData AS C  \
    INNER JOIN (   \
      SELECT Batch, Chip, max(Timestamp) as MaxDate   \
      FROM OrtelTe.dbo.ChirpSpecTrumData  \
      GROUP BY Batch, Chip   \
    ) AS CMaxDate   \
      ON C.Batch = CMaxDate.Batch AND C.Chip = CMaxDate.Chip AND C.Timestamp = CMaxDate.MaxDate   \
	INNER JOIN ( \
		SELECT Device_SN, BatchNo, ChipNo, max(TEST_DT) as MaxDate \
		FROM OrtelTE.dbo.ModuleDistortion \
		GROUP BY Device_SN, BatchNo, ChipNo \
	) AS Laser \
	ON C.Batch = Laser.BatchNo AND C.Chip = Laser.ChipNo \
	WHERE C.Batch != '1' AND C.Chip != '1' "
#%%
# !! Do not rerun if not needed, these queries take ~2 min
res3 = sql.ExecQuery(query3)
df3 = pd.DataFrame(data=res3, columns=['DeltaDate'])

res4 = sql.ExecQuery(query4)
df4 = pd.DataFrame(data=res4, columns=['DeltaDate'])  
df4 = df4[df4['DeltaDate'] > -1]

#%%
fig = plt.figure(figsize=(15, 6))
plt.hist(df3['DeltaDate'], bins=int(df3['DeltaDate'].max()/10) )
plt.title('Histogram of Time in Inventory for COBs Not Selected')
plt.savefig('Figures\\Histogram_All_COBs_Not_Selected.png')

#%%
fig = plt.figure(figsize=(15, 6))
plt.hist(df4['DeltaDate'], bins=int(df4['DeltaDate'].max()/10) )
plt.title('Histogram of Time Taken for COB to be Selected')
plt.savefig('Figures\\Histogram_All_COBs_Selected.png')
