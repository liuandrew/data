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
