#%%
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# ----------------------------------------
# LM to COB data correlations
# ----------------------------------------

#%%
# Load Data
link_df = pd.read_csv('LM_COB_link.csv')
cob_df = pd.read_csv('COB_data.csv')
lm_tds_df = pd.read_csv('Test Data Sheet for 1752A-C21 Shipped Units.csv')

# Link COB data to LM serials
link_df = pd.merge(link_df, cob_df, how='inner', on=['Batch', 'Chip'])
link_df['Device_SN'] = link_df['SN']
link_df = link_df.drop('SN', axis=1)

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
merged_df = pd.merge(link_df, lm_tds_df, how='inner', on='Device_SN')

X, y = prepare_data(merged_df, 'SMSR (dB)')

