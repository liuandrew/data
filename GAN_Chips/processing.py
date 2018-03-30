# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import numpy as np
import pandas as pd

crossover_dates = ['06/01/16', '07/01/16', '08/01/16', '09/01/16', '10/01/16',
                   '11/01/16', '01/01/17', '06/01/17']
csv_order = ['1116-6216', '6116-7216', '7116-8216', '8116-9216', '9116-10216',
             '10116-11216', '11116-1117', '123116-6217', '6117-123117']


df1 = pd.read_csv('6116-7216.csv')
columns = df1.columns

df = pd.DataFrame(columns = columns)

for (index, csv_prefix) in enumerate(csv_order):
    csv = csv_prefix + '.csv'
    df_temp = pd.read_csv(csv)
    #remove crossover dates
    if(index != 0):
        print(crossover_dates[index - 1])
        print('Number of chips found for crossover date in cumulative: ' + 
              str((df[df.TestDate == crossover_dates[index - 1]]).shape[0]))
        print('Number of chips found for crossover date in temp: ' + 
              str((df[df.TestDate == crossover_dates[index-1]]).shape[0]))
        df_temp = df_temp[df_temp.TestDate != crossover_dates[index - 1]]
    df = df.append(df_temp, ignore_index = True)


not_passed = df1[df1.PFcode != '0']

unwanter_wafer_no = ['DEBUG-1', 'DEBUG-2', 'DEBUG-4', 'DEBUG-5', 'ECL-3',
       'ECL-5', 'ELASER', 'GR&R', 'L', 'S', 'SP', 'SPC', 'SPC-1', 'SPC-2',
       'SPC-3', 'SPC-4', 'SPC-5', 'TESTER', 'TESTER-2', 'TESTER-3', 'STD', 'STD-2',
       'A', 'CORRE', 'HTC', 'LBM', 'SORRE', 'SPS', 'TESTER6', 'DEBUG', 'DEUG', 'GRR', 
       'GRR2', 'SPC1', 'SPC2', 'SPC3', 'SPC4', 'SPC5', 'SPC6', 'TEST', 'Q']

for i in unwanter_wafer_no:
    not_passed = not_passed[not_passed['WaferNo'] != i]
    
not_passed['WaferNo'] = not_passed['WaferNo'].apply(str)
not_passed['BatchNo'] = not_passed['BatchNo'].apply(str)
not_passed['ChipID'] = not_passed['ChipID'].apply(str)


def get_repeated_df(original):
    c2 = columns.append(pd.Index(['WLRange', 'IthRange', 'SpecChange', 'LargeIthChange', 'Warning']))
    repeated_df = pd.DataFrame(columns=c2)

    total_lines_read = 0
    next_checkpoint = 250

    #repeated_test_indices = []
    #repeated_test_id = []
    #repeated_ith_range = []
    #repeated_wl_range = []
    for waferno in original['WaferNo'].unique():
        on_wafer = original[original['WaferNo'] == waferno]
        for batchno in on_wafer['BatchNo'].unique():
            on_batch = on_wafer[on_wafer['BatchNo'] == batchno]
            for chipid in on_batch['ChipID'].unique():
                chips = on_batch[on_batch['ChipID'] == chipid]
                if(chips.shape[0] > 1):
                    chips = on_batch[on_batch['ChipID'] == chipid].copy()
                    #repeated found
                    #repeated_test_indices.append(chips.index[0])
                    #repeated_test_id.append((waferno, batchno, chipid))
                    #repeated_ith_range.append(chips.Ith.max() - chips.Ith.min())
                    #repeated_wl_range.append(chips.ModeECLwave.max() - chips.ModeECLwave.min())
                    #print('Repeated test found for: ', str(waferno), ',', str(batchno),
                    #      ',', str(chipid))
                    #print('Number found: ', chips.shape[0])
                    #print('Ith range: ', chips.Ith.max(), '-', chips.Ith.min())
                    #print('WL range: ', chips.ModeECLwave.max(), '-', chips.ModeECLwave.min())
                    ith_range = chips.Ith.max() - chips.Ith.min()
                    wl_range = chips.ModeECLwave.max() - chips.ModeECLwave.min()
                    large_ith = (ith_range >= 2.0)
                    wl_min_in_spec = (chips.ModeECLwave.min() >= 1555.0 and
                                      chips.ModeECLwave.min() <= 1585.0)
                    wl_max_in_spec = (chips.ModeECLwave.max() >= 1555.0 and
                                      chips.ModeECLwave.max() <= 1585.0)
                    wl_spec_change = (wl_min_in_spec != wl_max_in_spec)
                    warning = (wl_spec_change != large_ith)
                    chips.loc[:, 'WLRange'] = wl_range
                    chips.loc[:, 'IthRange'] = ith_range
                    chips.loc[:, 'SpecChange'] = wl_spec_change
                    chips.loc[:, 'LargeIthChange'] = large_ith
                    chips.loc[:, 'Warning'] = warning
                    repeated_df = repeated_df.append(chips)
                    
                total_lines_read += chips.shape[0]
                if(total_lines_read > next_checkpoint):
                    print('Lines read: ' + str(total_lines_read))
                    next_checkpoint += 250
    return repeated_df
    
def get_uniques(original):
    repeated_test_id = []
    for waferno in original['WaferNo'].unique():
        on_wafer = original[original['WaferNo'] == waferno]
        for batchno in on_wafer['BatchNo'].unique():
            on_batch = on_wafer[on_wafer['BatchNo'] == batchno]
            for chipid in on_batch['ChipID'].unique():
                repeated_test_id.append([waferno, batchno, chipid])
                
    return repeated_test_id