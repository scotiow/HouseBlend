#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
A coffee roulette app that pairs people up in an engineered serendipitous (pseudo-random) way
where repeat meetings are discouraged.

Key app terms:
- Period: A discrete time period in which meetings can occur. E.g. a week.
- Meeting: A meeting between two people occurring in a given period.
- Hansard: The settings file. It contains all input data required to run the scheduling algorithm and contains the historic schedule.
- Participant: An individual who is part of the meeting.
- Assistant: An individual who supports a participant, e.g. their PA.
- Parliament: The overall coffee roulette instance / set of participants.

Copyright (c) 2025 Scot Wheeler

This file is part of HouseBlend, which is licensed under the MIT License.
You may obtain a copy of the License at
https://opensource.org/licenses/MIT

"""

__author__ = "Scot Wheeler"
__license__ = "MIT"
__version__ = "0.4.0"

import numpy as np
import cvxpy as cp
import pandas as pd
import datetime as dt
import os
from faker import Faker
import warnings
import logging
logger = logging.getLogger(__name__)



def min_periods(n_people):
    """
    Calculate the minimum number of periods required that allows everyone to meet everyone else.

    Parameters
    ----------
    n_people : int
        The number of participants.

    Returns
    -------
    min_periods : int
        The minimum number of periods required that allows everyone to meet everyone else.

    """
    min_periods = n_people - (n_people % 2 == 0)  # the minimum number of periods required for everyone to meet everyone else
    return min_periods



def test_n_people(n_people):
    """ Generate the participant dataframe of n randomly named participants and associated assistants for testing purposes."""
    # create empty list to hold names
    participant_names = []
    
    # use the Faker library and append names to list
    fake = Faker()
    for _ in range(n_people):
        participant_names.append(fake.first_name() + " " + fake.last_name())
    
    # create the contacts dataframe
    participants_df = pd.DataFrame({
        "Person": participant_names,
        "email": ["{}@email.com".format(name.replace(" ", "")) for name in participant_names],
        "Assistant": participant_names,
        "Assistant email": ["{}@email.com".format(name.replace(" ", "")) for name in participant_names],
    }).set_index("Person")

    return participants_df


def create_initial_hansard(n_participants, parliament_name="example",
                           folderpath=None, filename=None, n_periods=None):
    """
    Create the initial hansard file for a new parliament.
    Warning, this will overwrite existing participant input data.
    """
    if folderpath is None:
        folderpath = parliament_name

    # create folder if doesn't already exist
    os.makedirs(folderpath, exist_ok=True)

    # set filename
    if filename is None:
        filename = '{}_hansard.xlsx'.format(parliament_name)
    # add xlsx if not already present
    if not filename.endswith('.xlsx'):
        filename += '.xlsx'

    filepath = os.path.abspath(os.path.join(folderpath, filename))

    if os.path.exists(filepath):
        print("Deleting previous hansard file")
        os.remove(filepath)

    # if you specify a test number, it creates a set of n random strings
    print("Creating new participant directory of {} length for testing at {}".format(n_participants, filepath))
    contacts = test_n_people(n_participants)

    # dates associated with periods
    if isinstance(n_periods, type(None)):
        n_periods = min_periods(n_participants)

    start_date = (dt.datetime.today() + dt.timedelta(days=(7 - dt.datetime.today().weekday()))).date()
    period_dates = pd.DataFrame({"Start Date": pd.date_range(start_date, periods=n_periods, freq='2W').values,
                                 "End Date": pd.date_range(start_date + dt.timedelta(days=14), periods=n_periods, freq='2W').values},
                                index=list(range(1, n_periods + 1)))
    period_dates.index.name = "Period"

    availability = pd.DataFrame(np.ones((contacts.shape[0], n_periods)), columns=["Period {}".format(i) for i in range(1, n_periods + 1)], index=contacts.index)

    schedule = pd.DataFrame("", index=contacts.index, columns=["Period {}".format(x) for x in range(1, n_periods + 1)])

    save_hansard(contacts, period_dates, availability, schedule,
                 parliament_name=parliament_name,
                 folderpath=folderpath,
                 filename=filename)

    # # create availability_all dataframe
    # availability_all = generate_availability_all(availability, contacts, period_dates)

    return contacts, period_dates, availability, schedule, generate_boolean_schedule(schedule)

# def generate_availability_all(availability, contacts, period_dates):
#     """
#     Generate availability_all dataframe from availability dataframe.
#     """
#     availability_all = pd.DataFrame(np.ones((contacts.shape[0], period_dates.shape[0])),
#                                     index=contacts.index,
#                                     columns=period_dates.index)
#     unavail_mask = availability == 0

#     # Ensure we use the same column labels as availability_all. availability may use
#     # "Period N" strings so try to map those to numeric period indices.
#     if not all(col in availability_all.columns for col in unavail_mask.columns):
#         try:
#             mapped_cols = [int(str(col).split()[-1]) for col in unavail_mask.columns]
#         except Exception:
#             mapped_cols = list(unavail_mask.columns)
#     else:
#         mapped_cols = list(unavail_mask.columns)

#     mask_df = pd.DataFrame(~unavail_mask.values, index=unavail_mask.index, columns=mapped_cols)
#     availability_all.loc[unavail_mask.index, mapped_cols] = availability_all.loc[unavail_mask.index, mapped_cols].where(mask_df, other=0)
#     return availability_all

def save_hansard(contacts, dates, availability, schedule, parliament_name="example", folderpath=None, filename=None):
    """
    Save the hansard file for a parliament.
    """
    if folderpath is None:
        folderpath = parliament_name

    # create folder if doesn't already exist
    os.makedirs(folderpath, exist_ok=True)

    # set filename
    if filename is None:
        filename = '{}_hansard.xlsx'.format(parliament_name)
    # add xlsx if not already present
    if not filename.endswith('.xlsx'):
        filename += '.xlsx'

    filepath = os.path.join(folderpath, filename)

    # if file does not exist, create it
    if not os.path.exists(filepath):
        with pd.ExcelWriter(filepath) as writer:
            contacts.to_excel(writer, sheet_name='Participants')
            dates.to_excel(writer, sheet_name='Dates')
            availability.to_excel(writer, sheet_name='Availability')
            schedule.to_excel(writer, sheet_name='Schedule')
    else:
        with pd.ExcelWriter(filepath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            contacts.to_excel(writer, sheet_name='Participants')
            dates.to_excel(writer, sheet_name='Dates')
            availability.to_excel(writer, sheet_name='Availability')
            schedule.to_excel(writer, sheet_name='Schedule')

def import_hansard(folderpath=None, filename=None, parliament_name="example",
                       test=False, n_periods=None):
    """
    Import participant input dataframe from excel if exits.

    Create testing version if test set to True or int.
    """
    if test is not False:
        return create_initial_hansard(n_participants=test if isinstance(test, int) else 4,
                                      parliament_name=parliament_name,
                                      folderpath=folderpath,
                                      filename=filename,
                                      n_periods=n_periods)

    if folderpath is None:
        folderpath = parliament_name

    # set filename
    if filename is None:
        filename = '{}_hansard.xlsx'.format(parliament_name)
    # add xlsx if not already present
    if not filename.endswith('.xlsx'):
        filename += '.xlsx'

    filepath = os.path.join(folderpath, filename)

    if os.path.exists(filepath):
        logger.info("Importing contacts")

        contacts = pd.read_excel(
            filepath,
            index_col=0,
            sheet_name="Participants"
        )

        dates = pd.read_excel(
            filepath,
            index_col=0,
            sheet_name="Dates"
        )

        availability = pd.read_excel(
            filepath,
            index_col=0,
            sheet_name="Availability"
        )
        # fill any nans in availability with 1s, assuming that blank cells mean available
        availability = availability.fillna(1)

        # availability_all = generate_availability_all(availability, contacts, dates)

        schedule = pd.read_excel(
            filepath,
            index_col=0,
            sheet_name="Schedule"
        )

        bool_schedule = generate_boolean_schedule(schedule)
    
        return contacts, dates, availability, schedule, bool_schedule
    else:
        print("No hansard file found, you must provide a list of persons taking part here: {}".format(filepath))
        return None


def create_mailto_link(meeting, subject=None):
    """
    Create a mailto link for a meeting. 
    
    Meeting is a series containing the fields:
    - Person 1 email
    - Person 2 email
    - Person 1 assistant email
    - Person 2 assistant email

    """
    to = f"{meeting['Person 1 assistant email']};{meeting['Person 2 assistant email']}"
    cc = f"{meeting['Person 1 email']};{meeting['Person 2 email']}"
    link = f"mailto:{to}?cc={cc}"
    if subject:
        link += f"&subject={subject.replace(' ', '%20')}"
    return link


def get_first_name(name):
    """ Extract the first name (the first word before the first space) """
    if isinstance(name, str) and name.strip():
        return name.strip().split()[0]
    return ""

def format_date_with_suffix(date):
    """ Format date string with format '28th May 2025' """
    # Ensure it's a datetime object
    if pd.isnull(date):
        return None
    date = pd.to_datetime(date)
    day = date.day
    # Get suffix
    if 10 <= day % 100 <= 20:
        suffix = 'th'
    else:
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
    return date.strftime(f"%A {day}{suffix} %B %Y")


def export_for_mailmerge2(contacts, bool_schedule, period, dates,
                          save=False, folderpath=None, parliament_name='example', save_prefix="HouseBlend", subject=None):
    """
    Export data for a particular period to create two separate mail merge dataframes.
    
    First dataframe: One row per person in each meeting (participants)
    Second dataframe: One row per assistant in each meeting (assistants)
    
    Parameters
    ----------
    contacts : DataFrame
        Contacts dataframe with person information
    bool_schedule : ndarray
        Boolean schedule array
    period : int
        Period number
    dates : DataFrame
        Dates dataframe
    save : bool, optional
        Whether to save to Excel files. Default is False.
    folderpath : str, optional
        Folder path to save files. Default is None.
    save_prefix : str, optional
        Prefix for save filenames. Default is "HouseBlend".
    subject : str, optional
        Subject for mailto links. Default is None.
        
    Returns
    -------
    tuple
        (participants_df, assistants_df) - Two dataframes for mail merge
    """
    
    # row and column index of each meeting for period. Indexes correspond to list participants.
    idxs, jdxs = np.where(bool_schedule[:, :, period - 1] == 1)
    persons1 = contacts.index[idxs]
    persons2 = contacts.index[jdxs]
    
    # Create base meetings dataframe
    meetings_base = pd.DataFrame({
        'Person 1': persons1, 
        'Person 2': persons2,
        'Person 1 email': contacts.loc[persons1, "email"].values,
        'Person 2 email': contacts.loc[persons2, "email"].values,
        'Person 1 assistant': contacts.loc[persons1, "Assistant"].values,
        'Person 2 assistant': contacts.loc[persons2, "Assistant"].values,
        'Person 1 assistant email': contacts.loc[persons1, "Assistant email"].values,
        'Person 2 assistant email': contacts.loc[persons2, "Assistant email"].values,
        'Start Date': dates.loc[period, "Start Date"].strftime('%d/%m/%Y'),
        'End Date': dates.loc[period, "End Date"].strftime('%d/%m/%Y'),
        'Start Date Long': format_date_with_suffix(dates.loc[period, "Start Date"]),
        'End Date Long': format_date_with_suffix(dates.loc[period, "End Date"])
    })
    
    # First dataframe: Participants mail merge
    participants_rows = []
    for _, meeting in meetings_base.iterrows():
        # Row for Person 1
        participants_rows.append({
            'mailto full name': meeting['Person 1'],
            'mailto first name': get_first_name(meeting['Person 1']),
            'mailto assistant full name': meeting['Person 1 assistant'],
            'paired with full name': meeting['Person 2'],
            'paired with first name': get_first_name(meeting['Person 2']),
            'start date': meeting['Start Date'],
            'end date': meeting['End Date'],
            'start date long': meeting['Start Date Long'],
            'end date long': meeting['End Date Long'],
            'mailto email': meeting['Person 1 email']
        })
        
        # Row for Person 2
        participants_rows.append({
            'mailto full name': meeting['Person 2'],
            'mailto first name': get_first_name(meeting['Person 2']),
            'mailto assistant full name': meeting['Person 2 assistant'],
            'paired with full name': meeting['Person 1'],
            'paired with first name': get_first_name(meeting['Person 1']),
            'start date': meeting['Start Date'],
            'end date': meeting['End Date'],
            'start date long': meeting['Start Date Long'],
            'end date long': meeting['End Date Long'],
            'mailto email': meeting['Person 2 email']
        })
    
    participants_df = pd.DataFrame(participants_rows)
    
    # Second dataframe: Assistants mail merge
    assistants_rows = []
    for _, meeting in meetings_base.iterrows():
        # Row for Person 1's assistant
        assistants_rows.append({
            'mailto assistant full name': meeting['Person 1 assistant'],
            'mailto assistant first name': get_first_name(meeting['Person 1 assistant']),
            'person full name': meeting['Person 1'],
            'paired with full name': meeting['Person 2'],
            'paired with assistant full name': meeting['Person 2 assistant'],
            'paired with assistant email': meeting['Person 2 assistant email'],
            'start date': meeting['Start Date'],
            'end date': meeting['End Date'],
            'start date long': meeting['Start Date Long'],
            'end date long': meeting['End Date Long'],
            'mailto email': meeting['Person 1 assistant email']
        })
        
        # Row for Person 2's assistant
        assistants_rows.append({
            'mailto assistant full name': meeting['Person 2 assistant'],
            'mailto assistant first name': get_first_name(meeting['Person 2 assistant']),
            'person full name': meeting['Person 2'],
            'paired with full name': meeting['Person 1'],
            'paired with assistant full name': meeting['Person 1 assistant'],
            'paired with assistant email': meeting['Person 1 assistant email'],
            'start date': meeting['Start Date'],
            'end date': meeting['End Date'],
            'start date long': meeting['Start Date Long'],
            'end date long': meeting['End Date Long'],
            'mailto email': meeting['Person 2 assistant email']
        })
    
    assistants_df = pd.DataFrame(assistants_rows)
    
    # Save to Excel files if requested
    if save:
        save_name = f'{parliament_name}_{save_prefix}_period_{period}_mailmerge.xlsx'
        
        if folderpath is None:
            folderpath = parliament_name
        save_path = os.path.join(folderpath, save_name)
        
        with pd.ExcelWriter(save_path) as writer:
            participants_df.to_excel(writer, sheet_name='Participants', index=False)
            assistants_df.to_excel(writer, sheet_name='Assistants', index=False)
    
    return participants_df, assistants_df


# def export_for_mailmerge(contacts, bool_schedule, period, dates,
#                         save=False, folderpath=None, save_prefix="HouseBlend", subject=None):
#     """
#     Export data for a particular period to be used in an Office based Mail Merge.
    
#     See README for description of mailmerge usage.

#     Parameters
#     ----------
#     contacts : TYPE
#         DESCRIPTION.
#     bool_schedule : TYPE
#         DESCRIPTION.
#     period : TYPE
#         DESCRIPTION.
#     dates : TYPE
#         DESCRIPTION.
#     save : TYPE, optional
#         DESCRIPTION. The default is False.
#     folderpath : TYPE, optional
#         DESCRIPTION. The default is None.
#     subject : TYPE, optional
#         DESCRIPTION. The default is None.

#     Returns
#     -------
#     None.

#     """
#     save_name = f'{save_prefix}_period_{period}_mailmerge.xlsx'
#     if folderpath is None:
#         save_path = os.path.join(parent_fullpath, save_name)
#     else:
#         save_path = os.path.join(folderpath, save_name)

#     # row and column index of each meeting for period. Indexes correspond to list participants.
#     idxs, jdxs = np.where(bool_schedule[:, :, period - 1] == 1)
#     persons1 = contacts.index[idxs]
#     persons2 = contacts.index[jdxs]

#     # core meetings dataframe
#     period_meetings = pd.DataFrame({'Person 1': persons1, 'Person 2': persons2,
#                                     'Person 1 email': contacts.loc[persons1, "email"].values,
#                                     'Person 2 email': contacts.loc[persons2, "email"].values,
#                                     'Person 1 assistant': contacts.loc[persons1, "Assistant"].values,
#                                     'Person 2 assistant': contacts.loc[persons2, "Assistant"].values,
#                                     'Person 1 assistant email': contacts.loc[persons1, "Assistant email"].values,
#                                     'Person 2 assistant email': contacts.loc[persons2, "Assistant email"].values,
#                                     # 'All emails': ["{}; {}".format(a1, a2) for a1, a2 in zip(contacts.loc[persons1, "Assistant email"].values, contacts.loc[persons2, "Assistant email"].values)],
#                                     'Start Date': dates.loc[period, "Start Date"].strftime('%d/%m/%Y'),
#                                     'End Date': dates.loc[period, "End Date"].strftime('%d/%m/%Y')
#                                     })
    
#     # mailto link
#     period_meetings["Mailto Link"] = period_meetings.apply(
#         lambda row: create_mailto_link(row, subject=subject),
#         axis=1
#         )

#     # create columns containing only first names
#     period_meetings["Person 1 first name"] = period_meetings['Person 1'].apply(
#         lambda name: get_first_name(name)
#     )
#     period_meetings["Person 2 first name"] = period_meetings['Person 2'].apply(
#         lambda name: get_first_name(name)
#     )
#     period_meetings["Person 1 assistant first name"] = period_meetings['Person 1 assistant'].apply(
#         lambda name: get_first_name(name)
#     )
#     period_meetings["Person 2 assistant first name"] = period_meetings['Person 2 assistant'].apply(
#         lambda name: get_first_name(name)
#     )

#     # create column containing all the emails of everyone involved in a meeting
#     period_meetings["All emails"] = period_meetings.apply(
#         lambda row: "; ".join([
#             f"<{row['Person 1 email']}>",
#             f"<{row['Person 2 email']}>",
#             f"<{row['Person 1 assistant email']}>",
#             f"<{row['Person 2 assistant email']}>"
#         ]),
#         axis=1
#     )

#     # Readable dates
#     period_meetings["Start Date Long"] = format_date_with_suffix(dates.loc[period, "Start Date"])
#     period_meetings["End Date Long"] = format_date_with_suffix(dates.loc[period, "End Date"])

#     ## expand dataframe so there is a row for every person involved in a meeting (i.e. 4 rows for every meeting)
#     ## this allows you to address each person individually which is a limitataion of office mailmerge.
#     # Select the columns you want to keep constant
#     base_columns = ['Person 1', 'Person 2', "Person 1 first name", "Person 2 first name",
#                     'Person 1 assistant', 'Person 2 assistant', "Person 1 assistant first name", "Person 2 assistant first name",
#                     'Start Date', 'End Date', "Start Date Long", "End Date Long",
#                     "All emails", "Mailto Link"]

#     # Melt the dataframe to turn the four email columns into a single 'mailto' column
#     melted = pd.melt(
#         period_meetings,
#         id_vars=base_columns,
#         value_vars=[
#             'Person 1 email',
#             'Person 2 email',
#             'Person 1 assistant email',
#             'Person 2 assistant email'
#         ],
#         var_name='Email Type',
#         value_name='mailto'
#     )

#     if save:
#         melted.to_excel(save_path)
#     return melted
    
    

def period_meeting_list(contacts, bool_schedule, period,
                        save=False, folderpath=None, parliament_name='example'):
    """
    Generate dataframe containing all pairs of people scheduled to meet in a given period.
    """


    idxs, jdxs = np.where(bool_schedule[:, :, period - 1] == 1)
    persons1 = contacts.index[idxs]
    persons2 = contacts.index[jdxs]

    period_meetings = pd.DataFrame({'Person 1': persons1, 'Person 2': persons2})
    
    if save:
        save_name = f'period_{period}_meeting_list.xlsx'
        if folderpath is None:
            folderpath = parliament_name
        save_path = os.path.join(folderpath, save_name)
        period_meetings.to_excel(save_path)

    return period_meetings


def generate_meeting_schedule(contacts, dates, availability, bool_schedule, save=False,
                              folderpath=None, parliament_name="example", filename=None):
    """
    Generate readable schedule from raw schedule numpy array.
    """

    periods = bool_schedule.shape[2]
    schedule = pd.DataFrame("", index=contacts.index, columns=["Period {}".format(x) for x in range(1, periods + 1)])
    for k in range(periods):
        paired_person = period_meeting_person(contacts, bool_schedule, k + 1, contacts.index.values)
        schedule.loc[:, "Period {}".format(k + 1)] = paired_person
    if save:
        # use the save hansard function
        save_hansard(contacts, dates, availability, schedule,
                     parliament_name=parliament_name, folderpath=folderpath, filename=filename)

    #     if folderpath is None:
    #         folderpath = parliament_name

    #     # set filename
    #     if filename is None:
    #         filename = '{}_hansard.xlsx'.format(parliament_name)
    #     # add xlsx if not already present
    #     if not filename.endswith('.xlsx'):
    #         filename += '.xlsx'

    #     filepath = os.path.join(folderpath, filename)
    #     with pd.ExcelWriter(filepath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    #         schedule.to_excel(writer, sheet_name='Schedule')
    return schedule


def period_meeting_person(contacts, bool_schedule, period, persons):
    """
    Return the person who is scheduled to meet with a given person in a given period.
    """
    if isinstance(persons, str):
        persons = [persons]
    paired_persons = []
    period_pairs = period_meeting_list(contacts, bool_schedule, period)
    for person in persons:
        mask = period_pairs == person
        if mask.sum().sum() == 0:
            paired_persons.append(np.nan)
        else:
            pair = period_pairs.loc[mask.sum(axis=1) == 1, :].to_numpy().flatten()
            paired_persons.append(pair[pair != person][0])
    return paired_persons if len(paired_persons) > 1 else paired_persons[0]


def penalty_weighting(difference, max_penalty, decay_rate=1):
    # what happens if there is no second meeting?
    return max_penalty * np.exp(-decay_rate * difference)


def participant_update(contacts, dates, availability, schedule, bool_schedule, parliament_name, folderpath=None, filename=None):
    """
    Handle any updates to the participant list within the participant sheet of the hansard.

    The participant list is treated as the master list. Any additions or removals are reflected in the schedule and availability.
    """
    if isinstance(schedule, type(None)):
        print("No schedule provided. Skipping participant update")
        return contacts, schedule, bool_schedule

    # check for differences in participants between contacts and schedule
    if (contacts.index.difference(schedule.index).shape[0] == 0) and (schedule.index.difference(contacts.index).shape[0] == 0):
        print("No participant changes")
        return contacts, dates, availability, schedule, bool_schedule
    else:
        # those missing from contacts need to be removed from bool schedule
        for_removal = schedule.index.difference(contacts.index)
        if for_removal.shape[0] > 0:
            print("Removing participants {}".format(for_removal.values))
            ixs = schedule.index.get_indexer(for_removal.values)
            bool_schedule = np.delete(bool_schedule, ixs, axis=0)
            bool_schedule = np.delete(bool_schedule, ixs, axis=1)

        # those missing from schedule need to be added to bool schedule
        for_addition = contacts.index.difference(schedule.index)
        if for_addition.shape[0] > 0:
            print("Adding participants {}".format(for_addition.values))
            for participant in for_addition:
                contacts_idx = contacts.index.get_indexer([participant])
                bool_schedule = np.insert(bool_schedule, contacts_idx[0], 0, axis=0)
                bool_schedule = np.insert(bool_schedule, contacts_idx[0], 0, axis=1)

        # update availability from comparison with contacts
        # remove rows from availability that are no longer in contacts
        availability = availability.loc[availability.index.intersection(contacts.index), :]

        # regenerate schedule dataframe
        schedule = generate_meeting_schedule(contacts, dates, availability, bool_schedule, parliament_name=parliament_name, folderpath=folderpath, filename=filename)
        
        # add rows to availability for new contacts, defaulting to available for all periods
        for participant in for_addition:
            contacts_idx = contacts.index.get_indexer([participant])
            availability = availability.append(pd.Series(1, index=availability.columns, name=participant))

        return contacts, dates, availability, schedule, bool_schedule

def check_dates_and_add(current_period, n_to_schedule, dates, availability, schedule, bool_schedule, parliament_name, folderpath=None, filename=None):
    """
    Check the dates sheet to ensure there are enough future datess for the number of periods we want to schedule.
    If not, add more dates based on the mode frequency (e.g. weekly, biweekly) of the existing dates.
    """
    # determine frequency of existing dates
    if dates.shape[0] > 1:
        freq = pd.infer_freq(dates["Start Date"])
        if freq is None:
            print("Could not infer frequency from existing dates. Defaulting to biweekly.")
            freq = '2W'
    else:
        print("Not enough existing dates to infer frequency. Defaulting to biweekly.")
        freq = '2W'
    
    for i in range(current_period, current_period + n_to_schedule):
        if i not in dates.index:
            print(f"Adding dates for period {i} as it is missing from the dates sheet.")
            # determine frequency of existing dates
            last_start_date = dates["Start Date"].max()
            last_end_date = dates["End Date"].max()
            new_start_date = last_start_date + pd.tseries.frequencies.to_offset(freq)
            new_end_date = last_end_date + pd.tseries.frequencies.to_offset(freq)
            new_dates = pd.DataFrame({"Start Date": [new_start_date], "End Date": [new_end_date]}, index=[i])
            dates = pd.concat([dates, new_dates])
    
    # add column names to availability and schedule if needed
    for i in range(current_period, current_period + n_to_schedule):
        if f"Period {i}" not in availability.columns:
            availability[f"Period {i}"] = 1  # default to available
        if f"Period {i}" not in schedule.columns:
            schedule[f"Period {i}"] = ""
    
    # add layers to bool_schedule if needed
    if bool_schedule.shape[2] < current_period + n_to_schedule - 1:
        additional_layers = current_period + n_to_schedule - 1 - bool_schedule.shape[2]
        bool_schedule = np.concatenate([bool_schedule, np.zeros((bool_schedule.shape[0], bool_schedule.shape[1], additional_layers))], axis=2)

    # return updated dates, schedule, and bool_schedule
    return dates, availability, schedule, bool_schedule


def generate_boolean_schedule(schedule):
    """
    Generates boolean schedule from imported schedule dataframe.
    """
    n_people = int(schedule.shape[0])
    n_periods = int(schedule.shape[1])
    bool_schedule = np.zeros((n_people, n_people, n_periods))
    for k in range(n_periods):
        idxs = schedule.index.get_indexer(schedule["Period {}".format(k + 1)].values)
        jdxs = np.where(idxs != -1)[0]  # ignore names that return nan for that period
        bool_schedule[idxs[idxs != -1], jdxs, k] = 1
        lower_mask = np.tril(np.ones((n_people, n_people)), k=0)
        bool_schedule[:, :, k] = bool_schedule[:, :, k] * lower_mask
    return bool_schedule

def run_schedule_optimisation2(contacts, dates, availability, schedule, bool_schedule, n_to_schedule,
                               current_period=1, verbose=False,
                               multiple_meetings='strict', save=False, folderpath=None,
                               parliament_name="example", filename=None):
    """
    Updated version of run schedule optimisation.
    """

    # note, current period is 1 indexed, but bool_schedule is 0 indexed, so need to subtract 1 when indexing bool_schedule
    
    # the number of people
    n_people = contacts.shape[0]

    # determine the number of periods ahead to schedule
    # if n_to_schedule is None, set to minimum required periods for everyone to meet everyone else
    if isinstance(n_to_schedule, type(None)):
        n_to_schedule = min_periods(n_people=n_people)  # the number periods (e.g. weeks if meetings are weekly) in the season

    # check dates and add if needed
    dates, availability, schedule, bool_schedule = check_dates_and_add(current_period, n_to_schedule, dates, availability, schedule, bool_schedule, parliament_name, folderpath=folderpath, filename=filename)

    # Define variable
    X = cp.Variable((n_people, n_people, n_to_schedule), boolean=True)

    # define constraints
    constraints = []

    # Get upper triangle mask (including diagonal)
    upper_mask = np.triu(np.ones((n_people, n_people)), k=0)

    for k in range(n_to_schedule):
        # upper triangle will always be 0
        constraints.append(cp.multiply(upper_mask, X[:, :, k]) == 0)

        # each person can only meet once per week, each person is represented by a row and column
        for i in range(n_people):
            constraints.append(cp.sum(X[i, :, k]) + cp.sum(X[:, i, k]) <= 1)

    # set unavailability
    # do this by setting the sum of the corresponding row and column to 0?
    for k, m in zip(range(n_to_schedule), range(current_period - 1, current_period - 1 + n_to_schedule)):
        period_unavail_idxs = availability.index.get_indexer(availability[availability[f'Period {m + 1}'] == 0].index)
        constraints.append(cp.sum(X[period_unavail_idxs, :, k]) == 0)
        constraints.append(cp.sum(X[:, period_unavail_idxs, k]) == 0)

    # backwards looking penalising repeat meetings
    # for each i, j pair, calculate the time since they last met
    weighting = np.zeros((n_people, n_people))
    for i in range(n_people):
        for j in range(i):
            if i != j:
                # Find the most recent period where i and j met
                last_meeting_period = 0
                for k in range(current_period - 1, -1, -1):
                    if (bool_schedule[i, j, k] == 1) or (bool_schedule[j, i, k] == 1): # check both i,j and j,i in case of any issues with upper triangle
                        last_meeting_period = max(last_meeting_period, k + 1)
            weighting[i, j] = 1/penalty_weighting(abs(current_period - last_meeting_period), max_penalty=1, decay_rate=0.1) 
    print(pd.DataFrame(weighting, index=contacts.index, columns=contacts.index))
    # forward looking penalising repeat meetings
    # only relevant if n_periods > 1
    if n_to_schedule > 1:
        # introducing a penalty for multiple meetings
        # To model X[i,j,k] * X[i,j,l] which is not DCP-compliant when X is boolean, you can introduce a new auxiliary binary variable Z[i,j,k,l] and enforce constraints that approximate this product:
        Z = cp.Variable((n_people, n_people, n_to_schedule, n_to_schedule), boolean=True)

        for i in range(n_people):
            for j in range(i):
                for k in range(n_to_schedule):
                    for L in range(k + 1, n_to_schedule):
                        # Enforce Z[i,j,k,l] = X[i,j,k] AND X[i,j,l]
                        # linear constraints to approximate the AND operation
                        constraints += [
                            Z[i, j, k, L] <= X[i, j, k], 
                            Z[i, j, k, L] <= X[i, j, L],
                            Z[i, j, k, L] >= X[i, j, k] + X[i, j, L] - 1,
                        ]

        penalty_terms = []

        for i in range(n_people):
            for j in range(i):
                for k in range(n_to_schedule):
                    for L in range(k + 1, n_to_schedule):
                        penalty = penalty_weighting(abs(k - L), max_penalty=1, decay_rate=0.1)
                        penalty_terms.append(penalty * Z[i, j, k, L])

        # ensure weighting same shape as X[:,:, :n_to_schedule] by duplicating rows and columns as needed
        weighting_expanded = np.zeros((n_people, n_people, n_to_schedule))
        for k in range(n_to_schedule):
            weighting_expanded[:, :, k] = weighting
        # objective function to include both past meetings penalty and future meetings penalty
        objective = cp.Maximize(cp.sum(cp.multiply(weighting_expanded, X[:, :, :])) - cp.sum(penalty_terms))

    else:
        weighting_expanded = np.zeros((n_people, n_people, n_to_schedule))
        for k in range(n_to_schedule):
            weighting_expanded[:, :, k] = weighting
        #  objective function to include past meetings penalty only
        print("Single period, only including past meeting penalty")
        objective = cp.Maximize(cp.sum(cp.multiply(weighting_expanded, X[:, :, :])))

    # Solve the problem
    # warnings.filterwarnings('ignore')

    problem = cp.Problem(objective, constraints)
    problem.solve(verbose=verbose)
    # to add a time limit, need to use solver specific options, therefore, need to be explicit as to what solver to use. 

    new_schedule = (X.value >= 0.5).astype(int)
    bool_schedule[:, :, current_period - 1:current_period - 1 + n_to_schedule] = new_schedule

    if save:
        # regenerate schedule dataframe
        schedule = generate_meeting_schedule(contacts, dates, availability, bool_schedule, parliament_name=parliament_name, folderpath=folderpath, filename=filename)

        # save updated hansard
        save_hansard(contacts, dates, availability, schedule,
                     parliament_name=parliament_name,
                     folderpath=folderpath,
                     filename=filename)

    return bool_schedule

    

def run_schedule_optimisation(contacts, dates, availability, schedule, bool_schedule, n_to_schedule,
                              current_period=1, verbose=False,
                              multiple_meetings='strict', save=False, folderpath=None,
                              parliament_name="example", filename=None):
    """
    Run the schedule optimisation using cvxpy.

    The method overview is:
    1) Define the decision variable X, a 3D boolean array where X[i,j,k] = 1 if person i meets person j in period k.
    2) Define constraints to ensure that:
        - No one meets themselves.
        - Each person meets at most one other person per period.
        - Participants are only scheduled in periods they are available.
        - Historical meetings (if any) are preserved.
    3) Define the objective function to maximise the total number of meetings while penalising repeat meetings based on the chosen strategy.
    4) Solve the optimisation problem using a suitable solver.
    5) Extract the optimal schedule from the decision variable.

    Todo:
    - decide if default is n_to_schedule = all remaining periods in season or 1 period
    - reduce size of decision variable by only including remaining periods to schedule, to do this can calculate a penalty matrix beforehand that accounts for historic meetings,\
          then multiply this by X in the objective function.




    Parameters
    ----------
    contacts : DataFrame
        Contacts dataframe with participant information.
    dates : DataFrame
        Dates dataframe with period dates.
    availability : DataFrame
        Availability dataframe indicating participant availability per period.
    schedule : DataFrame
        Schedule dataframe with current meeting schedule.
    bool_schedule : ndarray
        Boolean schedule array indicating current meeting schedule.
    n_to_schedule : int
        Number of periods to schedule.
    current_period : int, optional
        Current period number. The default is 1.
    verbose : bool, optional
        Whether to print solver output. The default is False.
    multiple_meetings : str, optional
        Strategy for handling multiple meetings. Options are 'strict', 'penalty', 'penaltytime'. The default is 'strict'.
    save : bool, optional
        Whether to save the updated schedule. The default is False.
    folderpath : str, optional
        Folder path to save the schedule. The default is None.
    parliament_name : str, optional
        Name of the parliament. The default is "example".
    filename : str, optional
        Filename to save the schedule. The default is None.

    Returns
    -------
    bool_schedule : ndarray
        Updated boolean schedule array.
    """
    # determine shape of schedule variable
    n_people = contacts.shape[0]
    
    # determine the number of periods ahead to schedule
    # if n_to_schedule is None, set to minimum required periods for everyone to meet everyone else
    if isinstance(n_to_schedule, type(None)):
        n_to_schedule = min_periods(n_people=n_people)  # the number periods (e.g. weeks if meetings are weekly) in the season
    # total periods is the number to schedule plus those already fixed
    total_periods = n_to_schedule + (current_period - 1)

    # Define variable
    X = cp.Variable((n_people, n_people, total_periods), boolean=True)

    # define constraints
    constraints = []

    # Get upper triangle mask (including diagonal)
    upper_mask = np.triu(np.ones((n_people, n_people)), k=0)

    for k in range(total_periods):
        # upper triangle will always be 0
        constraints.append(cp.multiply(upper_mask, X[:, :, k]) == 0)

        # each person can only meet once per week, each person is represented by a row and column
        for i in range(n_people):
            constraints.append(cp.sum(X[i, :, k]) + cp.sum(X[:, i, k]) <= 1)

    # set unavailability
    # do this by setting the sum of the corresponding row and column to 0?
    for k in range(total_periods):
        period_unavail_idxs = availability.index.get_indexer(availability[availability[f'Period {k + 1}'] == 0].index)
        constraints.append(cp.sum(X[period_unavail_idxs, :, k]) == 0)
        constraints.append(cp.sum(X[:, period_unavail_idxs, k]) == 0)

    # set historic periods if running part way through
    if current_period > 1:
        for k in range(current_period - 1):
            constraints.append(X[:, :, k] == bool_schedule[:, :, k])

    # =============================================================================
    #   Dealing with multiple meetings
    #   multiple_meetings = strict:
    #       only 1 meeting allowed
    # =============================================================================

    if multiple_meetings == "strict":
        # Each meeting between two persons can only happen at most once
        for i in range(n_people):
            for j in range(n_people):
                constraints.append(cp.sum(X[i, j, :]) <= 1)
        objective = cp.Maximize(cp.sum(X))

    elif multiple_meetings == 'penalty':
        pass
    
    elif multiple_meetings == "penaltytime":

        # =============================================================================
        #   Relaxation to allow multiple meetings
        # =============================================================================
        # only relevant if n_periods > 1
        if total_periods > 1:
            # introducing a penalty for multiple meetings
            # To model X[i,j,k] * X[i,j,l] which is not DCP-compliant when X is boolean, you can introduce a new auxiliary binary variable Z[i,j,k,l] and enforce constraints that approximate this product:
            Z = cp.Variable((n_people, n_people, total_periods, total_periods), boolean=True)

            for i in range(n_people):
                for j in range(i):
                    for k in range(total_periods):
                        for L in range(k + 1, total_periods):
                            # Enforce Z[i,j,k,l] = X[i,j,k] AND X[i,j,l]
                            # linear constraints to approximate the AND operation
                            constraints += [
                                Z[i, j, k, L] <= X[i, j, k], 
                                Z[i, j, k, L] <= X[i, j, L],
                                Z[i, j, k, L] >= X[i, j, k] + X[i, j, L] - 1,
                            ]

            penalty_terms = []

            for i in range(n_people):
                for j in range(i):
                    for k in range(total_periods):
                        for L in range(k + 1, total_periods):
                            penalty = penalty_weighting(abs(k - L), max_penalty=1, decay_rate=0.1)
                            penalty_terms.append(penalty * Z[i, j, k, L])

            # obective from V0.2.0
            # objective = cp.Maximize(cp.sum(X) - cp.sum(cp.hstack(penalty_terms)))

            # new objective for V0.3.0 that penalises there being no meetings
            objective = cp.Maximize(cp.sum(X) - cp.sum(cp.hstack(penalty_terms)) - total_periods * ((total_periods * n_people**2) - cp.sum(X)))

        else:
            objective = cp.Maximize(cp.sum(X))

    # Solve the problem
    warnings.filterwarnings('ignore')

    problem = cp.Problem(objective, constraints)
    problem.solve(verbose=verbose)
    # to add a time limit, need to use solver specific options, therefore, need to be explicit as to what solver to use. 

    bool_schedule = (X.value >= 0.5).astype(int)

    if save:
        # regenerate schedule dataframe
        schedule = generate_meeting_schedule(contacts, dates, availability, bool_schedule, parliament_name=parliament_name, folderpath=folderpath, filename=filename)

        # save updated hansard
        save_hansard(contacts, dates, availability, schedule,
                     parliament_name=parliament_name,
                     folderpath=folderpath,
                     filename=filename)

    return bool_schedule


if __name__ == "__main__":
    # SETUP
    # set below as per usage
    current_period = 1  # generating for this period and future periods. A value greater than 1 will set preceeding periods as fixed e.g. they have already happened.
    n_periods = 7  # set the horizon of the schedule (the term length). If set to None, this will be the minimum number of periods to ensure a meeting between every possible pair is arranged.
    # parent_dir = 'example'

    # # advanced - change if required
    # parent_fullpath = os.path.join('../', parent_dir)

    # # create directory if doesn't already exist
    # os.makedirs(parent_fullpath, exist_ok=True)

    # import contacts dataframe
    contacts, dates, availability, schedule, bool_schedule = import_hansard(test=4)

    bool_schedule = run_schedule_optimisation(contacts, dates, availability, schedule, bool_schedule,
                                               1, save=True)

    # meeting_schedule = generate_meeting_schedule(contacts, dates, availability, bool_schedule, save=True)

    # all_meetings_period = period_meeting_list(contacts, bool_schedule, 1, save=False)
    # all_meetings_period
