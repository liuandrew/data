from automation1.mssql import DatabaseManagerSQL
from automation1.utilities import *
import pandas as pd
import numpy as np
import os.path

#%%

# Connect to SQL
sql = DatabaseManagerSQL()

#%%

# ----------------------
# LM COB Linkage Data
# ----------------------

# Load LM Units data
df = pd.read_csv('Test Data Sheet for 1752A-C21 Shipped Units.csv')
# Convert SNs into csv string with each SN in single quotations
# first_hundred = df['Device_SN'].loc[0:99]
# first_hundred_csv = "'{}'".format(first_hundred.str.cat(sep="', '"))

# Load Data Checkpoints
checkpoint = read_json_file('LM_COB_link_checkpoint.json')
checkpoint_val = checkpoint['checkpoint']
if(os.path.isfile('LM_COB_link.csv')):
	LM_COB_link = pd.read_csv('LM_COB_link.csv')
else:
	LM_COB_link = pd.DataFrame()
#%%

# Loop data acquisition
for x in range(checkpoint_val, 500, 100):
	next_hundred = df['Device_SN'].loc[x:x+99]
	# Convert SNs into csv string with each SN in single quotations
	next_hundred_csv = "'{}'".format(next_hundred.str.cat(sep="', '"))

	# SQL Request
	lm_req = "SELECT TOP 100 t.BatchNo, t.ChipNo, t.TEST_DT, t.Device_SN \
	FROM OrtelTE.dbo.ModuleDistortion AS t \
	INNER JOIN ( \
		SELECT Device_SN, max(TEST_DT) AS MaxDate \
		FROM OrtelTE.dbo.ModuleDistortion \
		GROUP BY Device_SN \
	) AS tm \
	ON t.Device_SN = tm.Device_SN AND t.TEST_DT = tm.MaxDate \
	WHERE t.Device_SN IN ({})".format(next_hundred_csv)
	res = sql.ExecQuery(lm_req)

	# Append data to csv
	newdf = pd.DataFrame(data=res, columns=['Batch', 'Chip', 'Date', 'SN'])
	LM_COB_link = LM_COB_link.append(newdf, ignore_index=True)

	checkpoint['checkpoint'] = x+100
	write_json_file('LM_COB_link_checkpoint.json', checkpoint)
	LM_COB_link.to_csv('LM_COB_link.csv', index=False)
# end for

#%%
# -------------------
# COB Test Result Data
# -------------------

# Load COB links
# Load Data Checkpoints
checkpoint = read_json_file('COB_data_checkpoint.json')
checkpoint_val = checkpoint['checkpoint']

LM_COB_link = pd.read_csv('LM_COB_link.csv')

if(os.path.isfile('COB_data.csv')):
	COB_data = pd.read_csv('COB_data.csv')
else:
	COB_data = pd.DataFrame()

try:
    COB_data.Batch = COB_data.Batch.apply(str)
except:
    print('uninitialized')
#%%
# Loop data acquisition
for x in range(checkpoint_val, 200, 100):
	next_hundred_batch = LM_COB_link['Batch'].loc[x:x+99]
	next_hundred_batch = next_hundred_batch.apply(str)
	next_hundred_chip = LM_COB_link['Chip'].loc[x:x+99]
	# Convert SNs into csv string with each SN in single quotations
	next_hundred_batch_csv = "'{}'".format(next_hundred_batch.str.cat(sep="', '"))
	next_hundred_chip_csv = "'{}'".format(next_hundred_chip.str.cat(sep="', '"))

	# SQL Requests
	li_req = "SELECT TOP 200 LI.Batch, LI.ChipID, LI.Ith, LI.Power, LI.Efficiency, LI.WL, LI.KinkValue \
	FROM OrtelTE.dbo.LIDataBase AS LI \
		INNER JOIN (  \
			SELECT Batch, ChipID, max(TEST_DT) as MaxDate  \
			FROM OrtelTe.dbo.LIDataBase \
			GROUP BY Batch, ChipID  \
		) AS LIMaxDate  \
			ON LI.Batch = LIMaxDate.Batch AND LI.ChipID = LIMaxDate.ChipID AND LI.Test_DT = LIMaxDate.MaxDate  \
	WHERE LI.Batch IN ({batch}) AND LI.ChipID IN ({chip})".format(
		batch = next_hundred_batch_csv, chip = next_hundred_chip_csv)

	chirp_req = "SELECT TOP 200 C.Batch, C.Chip, C.TestTemp, C.RFTestFreq, C.PredictChannel, C.SMSR, C.Chirp, C.LISlope \
	FROM OrtelTE.dbo.ChirpSpecTrumData AS C \
		INNER JOIN (  \
			SELECT Batch, Chip, max(Timestamp) as MaxDate  \
			FROM OrtelTe.dbo.ChirpSpecTrumData \
			GROUP BY Batch, Chip  \
		) AS CMaxDate  \
			ON C.Batch = CMaxDate.Batch AND C.Chip = CMaxDate.Chip AND C.Timestamp = CMaxDate.MaxDate  \
	WHERE C.Batch IN ({batch}) AND C.Chip IN ({chip})".format(
		batch = next_hundred_batch_csv, chip = next_hundred_chip_csv)

	li_res = sql.ExecQuery(li_req)
	chirp_res = sql.ExecQuery(chirp_req)

	# Append data to csv
	li_df = pd.DataFrame(data=li_res, columns=['Batch', 'Chip', 'Ith', 'Power', 'Efficiency', 'WL', 'KinkValue'])
	chirp_df = pd.DataFrame(data=chirp_res, columns=['Batch', 'Chip', 'TestTemp', 'RFTestFreq', 'PredictChannel', 'SMSR', 'Chirp', 'LISlope'] )

	connected_df = pd.merge(li_df, chirp_df, on=['Batch', 'Chip'])
	COB_data = COB_data.append(connected_df, ignore_index=True)

	checkpoint['checkpoint'] = x+100
	write_json_file('COB_data_checkpoint.json', checkpoint)
	COB_data.to_csv('COB_data.csv', index=False)
# end for


#%%
# --------------------
# Link LM-COB to COB-data
# --------------------

COB_data = pd.read_csv('COB_data.csv')
LM_COB_link = pd.read_csv('LM_COB_link.csv')

COB_data.Batch = COB_data.Batch.apply(str)
LM_COB_link.Batch = LM_COB_link.Batch.apply(str)

LM_COB_data = pd.merge(LM_COB_link, COB_data, how='left', on=['Batch', 'Chip'])
LM_COB_data.to_csv('LM_COB_data.csv')