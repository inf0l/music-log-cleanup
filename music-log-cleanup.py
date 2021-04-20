#!/usr/bin/env python3
# Utility to parse and clean up xlsx exported from PGPTSession.

import pandas
from timecode import Timecode

readfile = pandas.read_excel(
    './EARTHRISE WILD RECOVERY-PG3 MUSIC LIST TRACK ORDER.xlsx')

df = pandas.DataFrame(readfile)
df.rename(columns={
    "Unnamed: 0": "TC IN",
    "Unnamed: 1": "TC OUT",
    "Unnamed: 5": "Track Name"
},
          inplace=True)

df1 = df[['Track Name', 'TC IN', 'TC OUT']]
df1.dropna(subset=["Track Name"], inplace=True)
df1.sort_values(['TC IN'], inplace=True)
df1.reset_index(drop=True, inplace=True)
df1['TC DURATION'] = df1.apply(
    lambda row: Timecode('25', row['TC OUT']) - Timecode('25', row['TC IN']),
    axis=1)

print(df1)
