#!/usr/bin/env python

"""
resolve2edl reads csv files exported from DaVinci Resolve and merges them
into a detailed EDL in csv format using pandas.
"""

__author__ = "Tal Zana"
__copyright__ = "Copyright 2020"
__license__ = "GPL"
__version__ = "0.1"

import pandas as pd
import os
from timecode import Timecode

#
# For debugging, set pandas display options
#

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.max_colwidth', 120)
pd.set_option('display.width', None)

#
# Filenames to read from current directory:
# MEDIAPOOL: csv file exported from Resolve using "Export Metadata from Selected Media Pool Clips"
# EDIT_INDEX: csv file expoted from Resolve using right-click on timeline, "Export Edit Index"

MEDIAPOOL = 'MediaPool.csv'
EDIT_INDEX = 'Montage.csv'

#
# Define the fields we want to keep in the csv files that Resolve creates
#

# Define the lists of columns we are interested in
# mp_full: all the columns that Resolve exports
# mp_keep: the ones we want to keep
mp_full = ['File Name', 'Clip Directory', 'Duration TC', 'Frame Rate', 'Audio Sample Rate',
           'Audio Channels', 'Resolution', 'Video Codec', 'Audio Codec', 'Reel Name',
           'Description', 'Comments', 'Keywords', 'Clip Color', 'Shot', 'Scene', 'Take',
           'Flags', 'Good Take', 'Shoot Day', 'Date Recorded', 'Camera #', 'Location',
           'Start TC', 'End TC', 'Start Frame', 'End Frame', 'Frames', 'Bit Depth',
           'Audio Bit Depth', 'Data Level', 'Date Modified', 'EDL Clip Name', 'Camera Type',
           'Camera Manufacturer', 'Shutter', 'ISO', 'Camera TC Type', 'Camera Firmware',
           'Lens Type', 'Lens Notes', 'Camera Aperture', 'Focal Point (mm)', 'Sound Roll #',
           'Reviewed By - DOP Reviewed']

mp_keep = ['File Name', 'Clip Directory', 'Duration TC', 'Frame Rate', 'Resolution',
           'Video Codec', 'Audio Codec', 'Description', 'Comments', 'Keywords',
           'Clip Color', 'Shot', 'Scene', 'Take', 'Flags', 'Camera #', 'Date Modified']
mp_drop = list(set(mp_full) - set(mp_keep))

# Do the same for the Edit Index
edit_full = ['#', 'Reel', 'Match', 'V', 'C', 'Dur', 'Source In', 'Source Out', 'Record In', 'Record Out',
             'Name', 'Comments', 'Source Start', 'Source End', 'Source Duration', 'Codec', 'Source FPS',
             'Resolution', 'Color', 'Notes', 'EDL Clip Name', 'Marker Keywords']

edit_keep = ['#', 'V', 'Source In', 'Source Out', 'Record In', 'Record Out', 'Name']
edit_drop = list(set(edit_full) - set(edit_keep))

# A list of clip names to exclude
excluded_clips = ['Fusion Title',
                  'Cross Fade 0 dB',
                  'Cross Dissolve',
                  'Audio Process Stream',
                  'Adjustment Clip',
                  'Dip To Color Dissolve',
                  'Solid Color']

# A list of track names to exclude
excluded_tracks = ['A11']

#
# Import the files and clean them up
#

# Import mediapool
mp = pd.read_csv(MEDIAPOOL, encoding='utf-16')

# Delete empty and unnecessary columns from mp
# (Resolve exports an empty column for some reason)
mp = mp.loc[:, ~mp.columns.str.contains('^Unnamed')]
mp.drop(mp_drop, axis=1, inplace=True)

# import edit index
edit = pd.read_csv(EDIT_INDEX, encoding='utf-8')

# Delete empty and unnecessary columns from edit
edit = edit.loc[:, ~edit.columns.str.contains('^Unnamed')]
edit.drop(edit_drop, axis=1, inplace=True)

# Delete rows with 'M2' which Resolve outputs for some reason
edit.drop(edit[edit['#'] == 'M2'].index, inplace=True)

#
# Resolve exports the Media Pool 'File Name' field with an extension,
# but the Edit Index 'Name' doesn't have an extension.
# We need to make mediapool.Filename = editindex.Name so we can merge the dataframes.
#
# Split the File Name column in mediapool into File Name and File Type
# os.path.splitext returns the tuple (name, ext)
# For the File Type we get rid of the first dot character (.jpg => jpg)
# .apply can't be used in place so we reassign the result to the File Name column
#

mp['File Type'] = mp['File Name'].apply(lambda x: os.path.splitext(x)[1][1:])
mp['File Name'] = mp['File Name'].apply(lambda x: os.path.splitext(x)[0])

# Rename the column in the mediapool
mp = mp.rename(columns={'File Name': 'Name'})

# Create the merged dataframe
df = pd.merge(edit, mp, on='Name', how='left')

#
# We can now do whatever we want with the data
#


def clips_without_source():
    # Query the dataframe for rows in which:
    # Name is not excluded above (dissolves etc),
    # Track is not excluded (commentary track),
    # Take is null (Take is used for the Source field in Resolve)
    # After the conditions we enter the columns we need
    return df.loc[~df['Name'].isin(excluded_clips) &
                  ~df['V'].isin(excluded_tracks) &
                  df['Take'].isnull(),
                  ['V', 'Name', 'File Type', 'Source In', 'Record In', 'Take']].sort_values('Record In')


def edl():
    # Query the dataframe for the full list except excluded clips and tracks
    return df.loc[~df['Name'].isin(excluded_clips) &
                  ~df['V'].isin(excluded_tracks)].sort_values('Record In')


# Write the EDL to a csv file in the current directory
edl().to_csv('edl.csv')
