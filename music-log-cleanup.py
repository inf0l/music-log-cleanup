#!/usr/bin/env python3
# Utility to parse and clean up xlsx exported from PGPTSession.

import pandas as pd
from timecode import Timecode
from pprint import pprint

readfile = pd.read_excel(
    './EARTHRISE WILD RECOVERY-PG3 MUSIC LIST TRACK ORDER.xlsx')

df = pd.DataFrame(readfile)

# Rename usable columns to something sane
df.rename(columns={
    "Unnamed: 0": "TC IN",
    "Unnamed: 1": "TC OUT",
    "Unnamed: 5": "Track Name"
},
          inplace=True)

# Create new dataframe with only usable columns
df1 = df[['Track Name', 'TC IN', 'TC OUT']]

# Drop empty rows and sort by TC
df1.dropna(subset=["Track Name"], inplace=True)
df1.sort_values(['TC IN'], inplace=True)
df1.reset_index(drop=True, inplace=True)

tracklist = df1.values.tolist()
tracks = [tracklist[0]]

# Clean up overlapping tracks
for currtrack in tracklist:
    prevtrack = tracks[-1]
    # Only check against first 10 characters of filename
    if prevtrack[0][:10] == currtrack[0][:10] and Timecode(
            '25', prevtrack[2]) >= Timecode('25', currtrack[1]):
        tracks.pop()
        tracks.append([prevtrack[0], prevtrack[1], currtrack[2]])
    else:
        tracks.append(currtrack)

df2 = pd.DataFrame(tracks, columns=list(["Track Name", "TC IN", "TC OUT"]))
df2.sort_values(['TC IN'], inplace=True)

# Calculate duration
df2['TC DURATION'] = df1.apply(
    lambda row: Timecode('25', row['TC OUT']) - Timecode('25', row['TC IN']),
    axis=1)

df2.to_excel("output.xlsx", index=False)
pprint(df2)
