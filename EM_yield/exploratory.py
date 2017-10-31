#%%
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime as dt

# ----------------------------------------
# LM to COB data correlations
# ----------------------------------------



#%%
# ------------------------------
# Function definitions
# ------------------------------

# Add suffix to COB data columns (not Batch, Chip, or Device_SN)
def suffix_cob_headers(df):
    static_headers = ['Batch', 'Chip', 'Device_SN']

    for column in df.columns:
        if(column not in static_headers):
            df['{0}_COB'.format(column)] = df[column]
            df = df.drop(column, axis=1)
    # end for
    return df
# end function


# Build a dataframe with all COB data, along with a single column from
# LM data to attempt to draw correlations from
def build_correlation_dataframe(df, target_column):
    ret_df = pd.DataFrame()

    # Extract all '_COB' suffixed features to be used in the dataframe
    for column in df.columns:
        if('_COB' in column):
            ret_df[column] = df[column]
    # end for

    if(target_column in df.columns):
        ret_df[target_column] = df[target_column]
    else:
        print('Target column does not exist in data')
        return None

    return ret_df
# end function


# Drop likely non-useful features 
def drop_non_useful_features(df):
    if('Date_COB' in df.columns):
        df = df.drop('Date_COB', axis=1)
    if('PredictChannel_COB' in df.columns):
        df= df.drop('PredictChannel_COB', axis=1)

    return df
# end function


# Drop values outside of certain quartile range which might be outliers
def filter_quantiles(df, column, low, high):
    quantiles = df.quantile([low, high])
    low_bound = quantiles[column].loc[low]
    high_bound = quantiles[column].loc[high]
    df = df[df[column] > low_bound]
    df = df[df[column] < high_bound]
    return df
# end function


# Split dataframe into X and y data sets
# Assume the last column is the dependent variable
def split_test_df(df):
    X = df.iloc[:, :-1]
    y = df.iloc[:, -1]
    return X, y
# end function


# Run all of the above functions to prepare X and y data sets
# for a given column from merged_df
def prepare_data(df, column):
    ret_df = build_correlation_dataframe(df, column)

    ret_df = drop_non_useful_features(ret_df)

    ret_df = ret_df.dropna()

    ret_df = filter_quantiles(ret_df, 'SMSR_COB', 0.01, 0.99)

    X, y = split_test_df(ret_df)

    return X, y
# end function

#%%
# Load Data
link_df = pd.read_csv('LM_COB_link.csv')
cob_df = pd.read_csv('COB_data.csv')
lm_tds_df = pd.read_csv('Test Data Sheet for 1752A-C21 Shipped Units.csv')
lm_tds_df = lm_tds_df.drop('Unnamed: 19', axis=1)
# Link COB data to LM serials
link_df = pd.merge(link_df, cob_df, how='inner', on=['Batch', 'Chip'])
link_df['Device_SN'] = link_df['SN']
link_df = link_df.drop('SN', axis=1)

# link_df = suffix_cob_headers(link_df)
merged_df = pd.merge(link_df, lm_tds_df, how='inner', on='Device_SN')

#%%
# --------------------
# DeltaDates
# --------------------
# Calculate the time between LM and COB tests to see how long the 'selection 
# time' for a given COB was, labelled as "DeltaDate" in the dataframe

#%%
df = pd.read_csv('LM_COB_data.csv')

df = df.dropna()
df = df.reset_index()

#%%
def delta_date( row ):
    lm_date = dt.strptime(row['Date'], '%Y-%m-%d %H:%M:%S')
    cob_date = dt.strptime(row['Timestamp'], '%Y-%m-%d %H:%M:%S')
    delta = pd.Series([(lm_date - cob_date).days])
    globals()['delta_date_col'] = globals()['delta_date_col'].append(delta)
    

delta_date_col = pd.Series()
df.apply(delta_date, axis=1)
delta_date_col = delta_date_col.reset_index()
delta_date_col = delta_date_col.drop('index', axis=1)
delta_date_col.columns = ['DeltaDate']

df['DeltaDate'] = delta_date_col['DeltaDate']
delta_date_df = df.drop('Unnamed: 0', axis=1)


#%%
# ---------------------
# Plot graphs
# ---------------------

cob_columns = ['Ith', 'Power', 'Efficiency',
       'WL', 'KinkValue', 'TestTemp', 'RFTestFreq', 'PredictChannel', 'SMSR',
       'Chirp', 'LISlope']

# These columns are either normally distributed or generally
# thought to be useful (i.e. SMSR), check histograms to see
cob_useful_columns = ['Ith', 'Power', 'Efficiency', 'SMSR', 
       'KinkValue', 'Chirp', 'LISlope']

# Remove outliers
# Drops to about 6k entries from 10k
for col in cob_useful_columns:
    delta_date_df = filter_quantiles(delta_date_df, col, 0.01, 0.99)
    merged_df = filter_quantiles(merged_df, col, 0.01, 0.99)

#%%
# Histograms for COB test data, just to see general distributions
fig = plt.figure(figsize=(15, 12))
for idx, col in enumerate(cob_columns):
    ax = plt.subplot(3, 4, idx+1)
    plt.tight_layout()
    plt.hist(delta_date_df[col])
    plt.title(col)
plt.savefig('Figures\\COB_Test_Data_Histograms.png')
#%%
# Histograms for 'useful columns', after removing outliers of <0.01, or >0.99
fig = plt.figure(figsize=(15, 8))
for idx, col in enumerate(cob_useful_columns):
    ax = plt.subplot(2, 4, idx+1)
    plt.tight_layout()
    plt.hist(delta_date_df[col])
    plt.title(col)
plt.savefig('Figures\\COB_Test_Data_Histograms_Removed_Outliers_Normal.png')

#%%
# -------------------
# Delta Date correlations
# -------------------

fig = plt.figure(figsize=(15, 8))
for idx, col in enumerate(cob_useful_columns):
    ax = plt.subplot(2, 4, idx+1)
    plt.tight_layout()
    plt.scatter(delta_date_df.head(1000)[col], delta_date_df.head(1000)['DeltaDate'])
    plt.title(col)
plt.savefig('Figures\\Test_Data_Versus_Delta_Date_First_1000.png')

#%%
# Histogram for Delta Dates
fig = plt.figure(figsize=(15, 6))
plt.hist(delta_date_df['DeltaDate'])
plt.title('Histograms for COB Time Spent in Inventory')
plt.savefig('Figures\\Delta_Date_Histogram.png')

fig = plt.figure(figsize=(15, 6))
plt.hist(delta_date_df['DeltaDate'], bins=30)
plt.title('Histograms for COB Time Spent in Inventory')
plt.savefig('Figures\\Delta_Date_Histogram_30_bins.png')

#%%
# -------------------------
# Test Data Correlations
# -------------------------

def plot_test_data_correlations( target_column ):
    fig = plt.figure(figsize=(15, 8))
    for idx, col in enumerate(cob_useful_columns):
        ax = plt.subplot(2, 4, idx+1)
        plt.tight_layout()
        plt.scatter(merged_df[col], merged_df[target_column])
        plt.title(col)
    save_name = target_column.replace('/', '-')
    plt.savefig('Figures\\COB_correlations\\COB_Test_Data_Correlation_With_{}.png'.format( 
        save_name ) )
    

def plot_test_data_correlations_sparse( target_column ):
    fig = plt.figure(figsize=(15, 8))
    for idx, col in enumerate(cob_useful_columns):
        ax = plt.subplot(2, 4, idx+1)
        plt.tight_layout()
        plt.scatter(merged_df.head(1000)[col], merged_df.head(1000)[target_column])
        plt.title(col)
    save_name = target_column.replace('/', '-')
    plt.savefig('Figures\\COB_correlations\\COB_Test_Data_Correlation_With_{}_First_1000.png'.format( 
        save_name ) )
    
#%%
# Run above two functions for each LM Test result
lm_columns = ['Iop (mA)', 'Ith (mA)', 'Laser Temperature (Top) (C)',
       'Slope Efficiency (mW/mA)', 'Optical Power (mW)', 'SMSR (dB)',
       'Tracking Error_MIN (dB)', 'Tracking Error_MAX (dB)',
       'Thermistor Resistance (Rth) (K-Ohm)', 'B Consttant (K)',
       'CNR @ 61.25 MHz (dB)', 'CNR @ 547.25 MHz (dB)', 'CSO (dB)', 'CTB (dB)',
       'Chirp (MHz/mW)', 'Monitor Current (micro-A)',
       'Frequency Response (dB)']

for col in lm_columns:
    plot_test_data_correlations(col)
    plot_test_data_correlations_sparse(col)
    
#%%
# Listed are potential columns that show hints of correlations between
# LM and COB test data
potential_lm_columns = ['Iop (mA)', 'Ith (mA)', 'Monitor Current (micro-A)', 
                        'Slope Efficiency (mW/mA)', 'SMSR (dB)' ]

