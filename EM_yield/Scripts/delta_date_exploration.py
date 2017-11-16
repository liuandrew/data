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

from sklearn.neighbors import KernelDensity
from sklearn.model_selection import GridSearchCV

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


#%%
# ------------------
# KDE Analysis
# ------------------

matched = pd.read_csv('../Data/Delta_Dates_Matched.csv')
unmatched = pd.read_csv('../Data/Delta_Dates_Unmatched.csv')

matched.drop('Unnamed: 0', axis=1, inplace=True)
unmatched.drop('Unnamed: 0', axis=1, inplace=True)

matched_kde = KernelDensity(kernel='gaussian', bandwidth = 9.0)
matched_kde.fit(matched)
unmatched_kde = KernelDensity(kernel='gaussian', bandwidth = 9.0)
unmatched_kde.fit(unmatched)

#%%

X_plot = pd.DataFrame(np.arange(0, 601, 1))
matched_log_dens = matched_kde.score_samples(X_plot)
matched_probabilities = np.exp(matched_log_dens)
unmatched_log_dens = unmatched_kde.score_samples(X_plot)
unmatched_probabilities = np.exp(unmatched_log_dens)

fig, ax = plt.subplots(figsize=(15, 10), nrows=2)
ax[0].plot(X_plot.loc[:, 0], matched_probabilities, '-', label='Gaussian Kernel (Matched)')
ax[1].plot(X_plot.loc[:, 0], unmatched_probabilities, '-', label='Gaussian Kernel (Unmatched)')
plt.show()

def chance_of_pick( day ):
	'''
	Pass a number of days unit has been in inventory
	Returns expected chance of being picked as decimal according to kde
	Intuitively seems incorrect, as at 600 days probability ~25%
	'''
	m = matched_probabilities[:day].sum() * matched.count()
	un = unmatched_probabilities[:day].sum() * unmatched.count()
	probability = m / (m + un)
	return probability

#%%
def counted_chance_of_pick( day ):
	'''
	Calculate chance of picked by num units beyond matched/unmatched
	'''
	greater_matched = matched[matched['DeltaDate'] > day].count()
	greater_unmatched = unmatched[unmatched['DeltaDate'] > day].count()
	probability = greater_matched / ( greater_matched + greater_unmatched )
	return probability

counted_probabilities = np.array([])
for i in range(601):
	counted_probabilities = np.append(counted_probabilities, counted_chance_of_pick(i))

counted_probabilities = pd.DataFrame(counted_probabilities, columns=['probability'])

fig, ax = plt.subplots(figsize=(15, 8))
ax.plot(counted_probabilities.index, 
				counted_probabilities['probability'], '-', label='Probability')
plt.title('Probability of Being Picked (Total Picked/Unpicked)')
plt.savefig('..\\Figures\\Probability_Of_Unit_Picked_Over_Time.png')
plt.show()


#%%
# --------------------------
# LM COB Correlations
# --------------------------
query1 = 'SELECT TOP 10000 C.RecordId, C.Timestamp, C.Batch, C.Chip, Laser.MaxDate as LaserTimestamp,  \
	DateDiff(d, C.Timestamp, Laser.MaxDate) as DeltaDate, PredictChannel,  \
	TestPwr Chip_TestPwr, LISlope Chip_LISlope, LIIth Chip_LIIth, PartNo Chip_PartNo,  \
	PeakWL Chip_PeakWL, DetuneAmpl Chip_DetuneAmpl, C.SMSR Chip_SMSR, \
	C.Chirp Chip_Chirp, ModCurrent Chip_ModCurrent, ModulationFreq Chip_ModulationFreq,  \
	RFpk Chip_RFpk, AlignPwr Chip_AlignPwr, Laser.RFClipping Laser_RFClipping, Laser.Ith Laser_Ith,  \
	Laser.Slopeff Laser_Slopeff, Laser.Temperature Laser_Temperature, Laser.Chirp Laser_Chirp,  \
	Laser.WaveLen Laser_WaveLen, Laser.SMSR Laser_SMSR, Laser.Ibb Laser_Ibb,  \
	Laser.PowerIbb Laser_PowerIbb, Laser.Iop Laser_Iop, Laser.Pop Laser_Pop,  \
	Laser.CSO_Chirp Laser_CSO_Chirp \
FROM OrtelTE.dbo.ChirpSpecTrumData AS C  \
	INNER JOIN (   \
		SELECT Batch, Chip, max(Timestamp) as MaxDate   \
		FROM OrtelTe.dbo.ChirpSpecTrumData  \
		GROUP BY Batch, Chip   \
	) AS CMaxDate   \
		ON C.Batch = CMaxDate.Batch AND C.Chip = CMaxDate.Chip AND C.Timestamp = CMaxDate.MaxDate   \
INNER JOIN ( \
	SELECT Device_SN, BatchNo, ChipNo, max(TEST_DT) as MaxDate, RFClipping, Ith, Slopeff, Temperature, Chirp, WaveLen, SMSR,  \
Ibb, PowerIbb, Iop, Pop, CSO_Chirp \
	FROM OrtelTE.dbo.ModuleDistortion \
	GROUP BY Device_SN, BatchNo, ChipNo, RFClipping, Ith, Slopeff, Temperature, Chirp, WaveLen, SMSR,  \
Ibb, PowerIbb, Iop, Pop, CSO_Chirp \
) AS Laser \
ON C.Batch = Laser.BatchNo AND C.Chip = Laser.ChipNo \
{} \
ORDER BY C.RecordId DESC'

# first loop
res1 = sql.ExecQuery(query1.format("WHERE C.Chip != '1' AND C.Batch != '1'".format(last_id)))
df1 = pd.DataFrame( res1, columns= [ 'RecordId', 'Timestamp', 'Batch', 'Chip', 
'LaserTimestamp', 'DeltaDate', 'PredictChannel', 'Chip_TestPwr', 
'Chip_LISlope', 'Chip_LIIth', 'Chip_PartNo', 'Chip_PeakWL', 'Chip_DetuneAmpl', 'Chip_SMSR',
'Chip_Chirp', 'Chip_ModCurrent', 'Chip_ModulationFreq', 'Chip_RFpk', 'Chip_AlignPwr', 
'Laser_RFClipping', 'Laser_Ith', 'Laser_Slopeff', 'Laser_Temperature', 'Laser_Chirp', 
'Laser_WaveLen', 'Laser_SMSR', 'Laser_Ibb', 'Laser_PowerIbb', 'Laser_Iop', 'Laser_Pop', 
'Laser_CSO_Chirp' ])
last_id = res1[-1][0]

looping = True

while(looping):
	res1 = sql.ExecQuery(query1.format("WHERE C.RecordId < {} AND C.Chip != '1' AND C.Batch != '1'".format(last_id)))
	df_res = pd.DataFrame( res1, columns= [ 'RecordId', 'Timestamp', 'Batch', 'Chip', 
	'LaserTimestamp', 'DeltaDate', 'PredictChannel', 'Chip_TestPwr', 
	'Chip_LISlope', 'Chip_LIIth', 'Chip_PartNo', 'Chip_PeakWL', 'Chip_DetuneAmpl', 'Chip_SMSR',
	'Chip_Chirp', 'Chip_ModCurrent', 'Chip_ModulationFreq', 'Chip_RFpk', 'Chip_AlignPwr', 
	'Laser_RFClipping', 'Laser_Ith', 'Laser_Slopeff', 'Laser_Temperature', 'Laser_Chirp', 
	'Laser_WaveLen', 'Laser_SMSR', 'Laser_Ibb', 'Laser_PowerIbb', 'Laser_Iop', 'Laser_Pop', 
	'Laser_CSO_Chirp' ])
	# get last record
	if(len(res1) < 1):
		looping = False
		print('Completed')
		break
	last_id = res1[-1][0]
	df1 = df1.append(df_res, ignore_index = True)
	df1.to_csv('temp.csv')	
	print('Loop completed, results found: {}'.format(len(res1)))
#%%
query2 = 'SELECT TOP 10000 C.RecordId, C.Timestamp, C.Batch, C.Chip, \
	(floor(DateDiff(d, C.Timestamp, CURRENT_TIMESTAMP) / 10.00) * 10) as DeltaDate, PredictChannel,  \
	TestPwr Chip_TestPwr, LISlope Chip_LISlope, LIIth Chip_LIIth, PartNo Chip_PartNo, \
	PeakWL Chip_PeakWL, DetuneAmpl Chip_DetuneAmpl, C.SMSR Chip_SMSR, \
	C.Chirp Chip_Chirp, ModCurrent Chip_ModCurrent, ModulationFreq Chip_ModulationFreq,  \
	RFpk Chip_RFpk, AlignPwr Chip_AlignPwr \
FROM OrtelTE.dbo.ChirpSpecTrumData AS C  \
	INNER JOIN (  \
		SELECT Batch, Chip, max(Timestamp) as MaxDate   \
		FROM OrtelTe.dbo.ChirpSpecTrumData  \
		GROUP BY Batch, Chip   \
	) AS CMaxDate   \
		ON C.Batch = CMaxDate.Batch AND C.Chip = CMaxDate.Chip AND C.Timestamp = CMaxDate.MaxDate   \
LEFT JOIN ( \
	SELECT Device_SN, BatchNo, ChipNo, max(TEST_DT) as MaxDate \
	FROM OrtelTE.dbo.ModuleDistortion \
	GROUP BY Device_SN, BatchNo, ChipNo ) AS Laser \
ON C.Batch = Laser.BatchNo AND C.Chip = Laser.ChipNo \
{} \
ORDER BY C.RecordId DESC'

# first loop
res2 = sql.ExecQuery(query2.format("WHERE C.Chip != '1' AND C.Batch != '1' AND Laser.Device_SN IS NULL"))
df2 = pd.DataFrame( res2, columns= [ 'RecordId', 'Timestamp', 'Batch', 'Chip', 
'DeltaDate', 'PredictChannel', 'Chip_TestPwr', 
'Chip_LISlope', 'Chip_LIIth', 'Chip_PartNo', 'Chip_PeakWL', 'Chip_DetuneAmpl', 'Chip_SMSR',
'Chip_Chirp', 'Chip_ModCurrent', 'Chip_ModulationFreq', 'Chip_RFpk', 'Chip_AlignPwr' ])
last_id = res2[-1][0]

looping = True

while(looping):
	res2 = sql.ExecQuery(query2.format("WHERE C.RecordId < {} AND C.Chip != '1' AND C.Batch != '1' \
		AND Laser.Device_SN IS NULL".format(last_id)))
	df_res = pd.DataFrame( res2, columns= [ 'RecordId', 'Timestamp', 'Batch', 'Chip', 
	'DeltaDate', 'PredictChannel', 'Chip_TestPwr', 
	'Chip_LISlope', 'Chip_LIIth', 'Chip_PartNo', 'Chip_PeakWL', 'Chip_DetuneAmpl', 'Chip_SMSR',
	'Chip_Chirp', 'Chip_ModCurrent', 'Chip_ModulationFreq', 'Chip_RFpk', 'Chip_AlignPwr' ])
	# get last record
	if(len(res2) < 1):
		looping = False
		print('Completed')
		break
	last_id = res2[-1][0]
	df2 = df2.append(df_res, ignore_index = True)
	df2.to_csv('temp.csv')	
	print('Loop completed, results found: {}'.format(len(res2)))

#%%
# Prepare linked and unlinked cob test data
df1 = pd.read_csv('..\\Data\\COB_LM_Linked_Data_2.csv')
df1 = df1.drop('Unnamed: 0', axis=1)
df1 = df1.drop_duplicates('RecordId')
df1 = df1.dropna()

identifier_cols = [ 'RecordId', 'Batch', 'Chip' ]
date_cols = [ 'LaserTimestamp', 'DeltaDate', 'Timestamp' ]
channel_col = [ 'PredictChannel' ]
chip_cols = [ 'Chip_TestPwr', 
	'Chip_LISlope', 'Chip_LIIth', 'Chip_PeakWL', 'Chip_SMSR',
	'Chip_Chirp', 'Chip_ModCurrent', 'Chip_ModulationFreq', 'Chip_RFpk', 
	'Chip_AlignPwr' ]
not_norm_chip_cols = [ 'Chip_DetuneAmpl', 'Chip_PartNo' ]
laser_cols = [ 'Laser_RFClipping', 'Laser_Ith', 'Laser_Slopeff', 'Laser_Temperature', 'Laser_Chirp', 
	'Laser_WaveLen', 'Laser_SMSR', 'Laser_Ibb', 'Laser_PowerIbb', 'Laser_Iop', 'Laser_Pop', 
	'Laser_CSO_Chirp' ]

def drop_outliers( df, column ):
	'''
	Remove outliers based on simple calculation of normal distribution and 
	5 std dev
	'''
	percentiles = np.percentile( df[ column ], [25, 50, 75] )
	mu = percentiles[1]
	sig = 0.74 * ( percentiles[2] - percentiles[0] )
	df = df[ df[ column ] > (mu - 5 * sig) ] 
	df = df[ df[ column ] < (mu + 5 * sig) ] 
	return df

for col in chip_cols:
    print( col )
    df1 = drop_outliers( df1, col )
    
df2 = pd.read_csv('..\\Data\\COB_Non_Linked_Data.csv')
df2 = df2.drop('Unnamed: 0', axis=1)
df2 = df2.drop_duplicates('RecordId')
df2 = df2.dropna()
for col in chip_cols:
    print( col )
    df2 = drop_outliers( df2, col )

#%%
# -------------------
# Naive Bayes Classifier
# -------------------
# Run prepare df1 and df2 linked and unlinked cob test data
from sklearn.naive_bayes import GaussianNB
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score
from sklearn.cross_validation import cross_val_score

X = pd.DataFrame( df1.loc[ :, chip_cols ] )
X = X.append( df2.loc[ :, chip_cols ], ignore_index = True )
y = pd.Series( np.ones( df1.shape[0] ) )
y = y.append( pd.Series( np.zeros( df2.shape[0] ) ), ignore_index = True )

# shuffle rows
X[ 'label' ] = y
X = X.sample( frac = 1 ).reset_index( drop = True )
y = X[ 'label' ]
X = X.drop( 'label', axis=1 )

''' X_train, X_test, y_train, y_test = train_test_split(X, y, test_size = 0.1)

model = GaussianNB()
model.fit( X_train, y_train )
accuracy_score( y_test, model.predict( X_test ) ) '''

# cross-val score
model = GaussianNB()
model.fit( X, y )
cross_val_score( model, X, y, cv=5 )