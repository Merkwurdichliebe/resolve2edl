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


class Resolve():
    def __init__(self, media, edit, output, fps=25):
        self.media_filename = media
        self.edit_filename = edit
        self.output_filename = output
        self.fps = fps

        # Flags and options
        self.include_idx_in_export = False
        self.exist_clips_with_null_source = False
        self.flag_duration_issue = False
        self.null_source_suffix = '-no-source'

        # Read the files and display basic info
        self.read_files()
        self.show_info()
        self.show_tracks()

        # Remove unneeded information from both dataframes
        self.prepare_edit()
        self.prepare_media()

        # Display additional information,
        # merge both dataframes and export to new file
        self.show_filetypes()
        self.join_tables()
        self.write_file()
        self.show_stats()
        self.cleanup()

    def read_files(self):
        self.media = pd.read_csv(self.media_filename, encoding="utf-16")
        self.edit = pd.read_csv(self.edit_filename, encoding="utf-8")

    def show_info(self):
        # Calculate timeline length from start and end TC
        timeline_start_tc = Timecode(self.fps, self.edit.iloc[0]['Record In'])
        timeline_end_tc = Timecode(self.fps, self.edit.iloc[-1]['Record Out'])
        timeline_duration_tc = timeline_end_tc - timeline_start_tc

        # Output to console
        self.print_title('timeline information')
        self.print_c('Frame rate', str(self.fps) + ' fps')
        self.print_c('Timeline start TC', timeline_start_tc)
        self.print_c('Timeline end TC', timeline_end_tc)
        self.print_c('Timeline duration', timeline_duration_tc)
        print('')

    def show_tracks(self):
        # Create a list of used video and audio tracks
        tracks = self.edit['V'].value_counts().index.to_list()
        v_tracks = [t for t in tracks if t.startswith('V')]
        a_tracks = [t for t in tracks if t.startswith('A')]

        # Natural sort: the function converts to integer
        # the portion of the string after the 1st character.
        v_tracks.sort(key=lambda x: int(x[1:]))
        a_tracks.sort(key=lambda x: int(x[1:]))

        # Output to console
        self.print_c(str(len(v_tracks)) + ' Video tracks', ' '.join(v_tracks))
        self.print_c(str(len(a_tracks)) + ' Audio tracks', ' '.join(a_tracks))

    def prepare_edit(self):
        e = self.edit

        # Keep only the columns we need
        e = e[EDIT_COLUMNS_TO_KEEP]

        # Drop rows with no values in 'Name'
        e = e[e['Name'].notna()]

        # Drop rows with values in 'Name' matching the list
        e = e[~e['Name'].str.contains(
            '|'.join(EDIT_ROWS_TO_IGNORE), na=False)]

        # Keep only the tracks we need
        e = e[~e['V'].isin(EDIT_TRACKS_TO_EXCLUDE)]

        self.edit = e

    def prepare_media(self):
        m = self.media

        # Keep only the columns we need
        m = m[MEDIA_COLUMNS_TO_KEEP]

        # Rename columns
        m = m.rename({
            'File Name': 'Name',
            'Take': 'Source',
            'Scene': 'Reference',
            'Camera #': 'Fonds'},
            axis='columns')

        # The Media Pool lists file names with extensions so we remove these,
        # making the file name the same as the "Name" column
        # in the edit dataframe
        m['Extension'] = m['Name'] \
            .apply(lambda x: os.path.splitext(x)[1]) \
            .str.replace('.', '', regex=False)
        m['Name'] = m['Name'] \
            .apply(lambda x: os.path.splitext(x)[0])

        self.media = m

    def join_tables(self):
        # Join the two dataframes based on the clip name
        edl = pd.merge(self.edit, self.media, on='Name', how='left')

        # Remove rows where Source is in the ignore list
        edl = edl[~edl['Source'].isin(MEDIA_SOURCES_TO_IGNORE)]

        # Create a separate dataframe for rows which have no 'Source' entry
        # and remove those rows from the main dataframe
        self.edl_no_source = edl[edl['Source'].isnull()]
        if len(self.edl_no_source) > 0:
            self.exist_clips_with_null_source = True
            # Drop rows with no values in 'Source'
            edl = edl[edl['Source'].notna()]

        self.print_title('clip durations')

        # Calculate a 'Duration' column
        edl['Duration'] = edl.apply(
            lambda row: self.get_tc_delta(row), axis='columns')

        if not self.flag_duration_issue:
            print('No issues')

        # Create a column to mark clips which are used on audio tracks
        edl['Track'] = edl['V'].apply(lambda x: 'AUDIO' if x[0] == 'A' else '')

        # Drop the index column
        edl = edl.reset_index(drop=True)

        # Dataframe should already be sorted by 'Record In'
        # but we make sure it is
        edl = edl.sort_values('Record In')

        self.edl = edl

    def write_file(self):
        # Write the EDL to an Excel file in the current directory
        self.edl.to_excel(self.output_filename + '.xlsx',
                          index=self.include_idx_in_export)

        # Output a separate file for clips with null sources
        if (self.exist_clips_with_null_source):
            f = self.output_filename + self.null_source_suffix + '.xlsx'
            self.edl_no_source.to_excel(f, index=self.include_idx_in_export)
            self.print_title('null sources')
            print(f'{len(self.edl_no_source)} clips with no source assigned')
            print(f'Exported to separate file: {f}')

        # Display message
        self.print_title('output')
        print(
            f"Merged Media Pool '{self.media_filename}' "
            f"and Edit Index '{self.edit_filename}' "
            f"to '{self.output_filename}.xlsx'.\n"
            f"(Total {len(self.edl)} clips)"
            )

    def show_stats(self):
        self.print_title('head and tail of edl')
        print(self.edl)
        self.print_title('edl stats')
        print(self.edl.describe())

    def cleanup(self):
        print('\nDone.')

    def show_filetypes(self):
        self.print_title('media pool file types')
        print(self.media['Extension'].value_counts().to_string())

    def get_tc_delta(self, row):
        record_in = Timecode(self.fps, row['Record In'])
        record_out = Timecode(self.fps, row['Record Out'])
        try:
            delta = (record_out-record_in)
        except ValueError:
            print(
                f'Timecode problem in clip: IN {record_in} OUT {record_out} '
                '(duration set to zero)'
                )
            delta = 0
            self.flag_duration_issue = True
        return delta

    @staticmethod
    def print_c(field, value):
        print(field.ljust(25) + str(value))

    @staticmethod
    def print_title(s):
        print('\n' + s.upper())
        print('-' * len(s) + '\n')


r = Resolve('MediaPool.csv', 'Montage.csv', 'edl')
