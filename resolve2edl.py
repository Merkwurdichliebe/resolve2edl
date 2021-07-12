#!/usr/bin/env python

"""
resolve2edl reads csv files exported from DaVinci Resolve
for both the Media Pool and a timeline's Edit Index
and merges them into a detailed EDL in Excel format using pandas.
"""

__author__ = "Tal Zana"
__copyright__ = "Copyright 2021"
__license__ = "GPL"
__version__ = "1.0"

import os
import pandas as pd
from timecode import Timecode

MEDIAPOOL_FILENAME = 'MediaPool.csv'
EDIT_INDEX_FILENAME = 'Montage.csv'
OUTPUT_FILENAME = 'edl'
NULL_SOURCE_SUFFIX = '-no-source'
INCLUDE_INDEX_IN_EXPORT = False
FPS = 25

exist_clips_with_null_source = False

EDIT_COLUMNS_TO_KEEP = [
    'Name',
    'Source In',
    'Source Out',
    'Record In',
    'Record Out',
    'V'
]

EDIT_ROWS_TO_IGNORE = [
    'Fusion Title',
    'Cross Fade 0 dB',
    'Cross Dissolve',
    'Audio Process Stream',
    'Adjustment Clip',
    'Dip To Color Dissolve',
    'Solid Color'
]

# We want to keep all audio tracks
EDIT_TRACKS_TO_EXCLUDE = [
    'V5',
    'V6',
    'V7',
    'V8',
    'V9',
    'V10',
    'V11',
    'V12'
]

MEDIA_COLUMNS_TO_KEEP = [
    'File Name',
    'Take',
    'Camera #',
    'Scene',
    'Comments',
    'Keywords'
]

MEDIA_SOURCES_TO_IGNORE = [
    'Sound FX Tal',
    'Tournage',
    'Musique Bibliothèque',
    'Musique Tal',
    'Musique sous droits'
]


def get_tc_delta(row):
    record_in = Timecode(FPS, row['Record In'])
    record_out = Timecode(FPS, row['Record Out'])
    try:
        delta = (record_out-record_in)
    except ValueError:
        print(
            f'Timecode problem in clip: IN {record_in} OUT {record_out} '
            '(duration set to zero)'
            )
        delta = 0
    return delta


print('\nCREATING EDL')
print('------------\n')

# ------------------------------------------------------------------------------
# Read the exported CSV files
# ------------------------------------------------------------------------------

media = pd.read_csv(MEDIAPOOL_FILENAME, encoding="utf-16")
edit = pd.read_csv(EDIT_INDEX_FILENAME, encoding="utf-8")

# ------------------------------------------------------------------------------
# Simplify the Edit dataframe
# ------------------------------------------------------------------------------

# Keep only the columns we need
edit = edit[EDIT_COLUMNS_TO_KEEP]

# Drop rows with no values in 'Name'
edit = edit[edit['Name'].notna()]

# Drop rows with values in 'Name' matching the list
edit = edit[~edit['Name'].str.contains(
    '|'.join(EDIT_ROWS_TO_IGNORE), na=False)]

# Keep only the tracks we need
edit = edit[~edit['V'].isin(EDIT_TRACKS_TO_EXCLUDE)]

# ------------------------------------------------------------------------------
# Simplify the Media dataframe
# ------------------------------------------------------------------------------

# Keep only the columns we need
media = media[MEDIA_COLUMNS_TO_KEEP]

# Rename columns
media = media.rename({
    'File Name': 'Name',
    'Take': 'Source',
    'Scene': 'Reference',
    'Camera #': 'Fonds'},
    axis='columns')

# The Media Pool lists file names with extensions so we remove these,
# making the file name the same as the "Name" column in the edit dataframe
media['Name'] = media['Name'].apply(lambda x: os.path.splitext(x)[0])

# ------------------------------------------------------------------------------
# Join the two dataframes and perform cleanup
# We need to remove rows with ignored sources *after* the join operation
# ------------------------------------------------------------------------------

# Join the two dataframes based on the clip name
df = pd.merge(edit, media, on='Name', how='left')

# Remove rows where Source is in the ignore list
df = df[~df['Source'].isin(MEDIA_SOURCES_TO_IGNORE)]

# Create a separate dataframe for rows which have no 'Source' entry
# and remove those rows from the main dataframe
df_no_source = df[df['Source'].isnull()]
if len(df_no_source) > 0:
    exist_clips_with_null_source = True
    # Drop rows with no values in 'Source'
    df = df[df['Source'].notna()]

# Calculate a 'Duration' column
df['Duration'] = df.apply(lambda row: get_tc_delta(row), axis='columns')

# Create a column to mark clips which are used on audio tracks
df['Track'] = df['V'].apply(lambda x: 'AUDIO' if x[0] == 'A' else '')

# Drop the index column
df = df.reset_index(drop=True)

# Dataframe should already be sorted by 'Record In'
# but we make sure it is
df = df.sort_values('Record In')

# Write the EDL to an Excel file in the current directory
df.to_excel(OUTPUT_FILENAME + '.xlsx',
            index=INCLUDE_INDEX_IN_EXPORT)

# Output a separate file for clips with null sources
if (exist_clips_with_null_source):
    f = OUTPUT_FILENAME + NULL_SOURCE_SUFFIX + '.xlsx'
    df_no_source.to_excel(f, index=INCLUDE_INDEX_IN_EXPORT)
    print('\nNULL SOURCES')
    print('--------------')
    print(f'\n{len(df_no_source)} clips with no source assigned')
    print(f'Exported to separate file: {f}')

# ------------------------------------------------------------------------------
# Console output
# ------------------------------------------------------------------------------

print('\nHEAD AND TAIL OF EDL')
print('--------------------')
print(df)
print('\nEDL STATS')
print('---------')
print(df.describe())
print('\nOUTPUT')
print('------\n')
# Display message
print(
    f"Merged Media Pool '{MEDIAPOOL_FILENAME}' "
    f"and Edit Index '{EDIT_INDEX_FILENAME}'\n"
    f"to '{OUTPUT_FILENAME}.csv' & '{OUTPUT_FILENAME}.xlsx'.\n"
    f"(Total {len(df)} clips)\n"
    f"Done.\n"
    )
