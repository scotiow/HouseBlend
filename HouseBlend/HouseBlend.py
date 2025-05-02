#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Apr 30 18:13:16 2025.

Copyright (c) 2025 Scot Wheeler

This file is part of HouseBlend, which is licensed under the MIT License.
You may obtain a copy of the License at
https://opensource.org/licenses/MIT

"""

__author__ = "Scot Wheeler"
__license__ = "MIT"
__version__ = "0.3.2"

import numpy as np
import cvxpy as cp
import pandas as pd
import datetime as dt
import os
from faker import Faker
import warnings


def min_periods(n_people):
    min_periods = n_people - (n_people % 2 == 0)  # the minimum number of periods required for everyone to meet everyone else
    return min_periods

def test_n_people(test):
    participant_names = []
    fake = Faker()
    for _ in range(test):
        participant_names.append(fake.first_name() + " " + fake.last_name())
    contacts = pd.DataFrame({
        "Person": participant_names,
        "email": ["{}@email.com".format(name.replace(" ", "")) for name in participant_names],
        "Assistant": participant_names,
        "Assistant email": ["{}@email.com".format(name.replace(" ", "")) for name in participant_names],
    }).set_index("Person")

    return contacts


def import_contacts_df(folderpath=None, filename='contacts.xlsx', test=False, n_periods=None):
    """
    Import contacts dataframe from excel if exits. Otherwise create testing version.
    """

    if test != False:
        return test_import_files(folderpath, filename, test, n_periods)

    if folderpath is None:
        filepath = os.path.join(parent_fullpath, filename)
    else:
        filepath = os.path.join(folderpath, filename)

    if os.path.exists(filepath):
        print("Importing contacts")

        contacts = pd.read_excel(filepath, index_col=0, sheet_name="Participants")

        dates = pd.read_excel(filepath, index_col=0, sheet_name="Dates")

        availability = pd.read_excel(filepath, index_col=0, sheet_name="Availability")
        availability_all = pd.DataFrame(np.ones((contacts.shape[0], dates.shape[0])), index=contacts.index, columns=dates.index)
        unavail_mask = availability == 0
        availability_all.loc[unavail_mask.index, unavail_mask.columns] = availability_all.loc[unavail_mask.index, unavail_mask.columns].where(~unavail_mask, other=0)

        return contacts, dates, availability_all
    else:
        print("No contacts found, you must provide a list of persons taking part here: {}".format(filepath))
        return None


def test_import_files(folderpath, filename, test, n_periods):

    if folderpath is None:
        filepath = os.path.join(parent_fullpath, filename)
    else:
        filepath = os.path.join(folderpath, filename)

    print("Deleting previous contacts file")
    if os.path.exists(filepath):
        os.remove(filepath)

    if test == True:
        test=5

    # if you specify a test number, it creates a set of n random strings
    print("Creating new contact directory of {} length for testing".format(test))
    contacts = test_n_people(test)

    # dates associated with periods
    if isinstance(n_periods, type(None)):
        n_periods = contacts.shape[0] - (contacts.shape[0] % 2 == 0)  # the number periods (e.g. weeks if meetings are weekly) in the season

    start_date = (dt.datetime.today() + dt.timedelta(days=(7 - dt.datetime.today().weekday()))).date()
    period_dates = pd.DataFrame({"Date": pd.date_range(start_date, periods=n_periods, freq='2W').values}, index=list(range(1, n_periods + 1)))
    period_dates.index.name = "Period"

    availability = pd.DataFrame({})

    with pd.ExcelWriter(filepath) as writer:
        contacts.to_excel(writer, sheet_name='Participants')
        period_dates.to_excel(writer, sheet_name='Dates')
        availability.to_excel(writer, sheet_name='Availability')

    availability_all = pd.DataFrame(np.ones((contacts.shape[0], period_dates.shape[0])), index=contacts.index, columns=period_dates.index)
    unavail_mask = availability == 0
    availability_all.loc[unavail_mask.index, unavail_mask.columns] = availability_all.loc[unavail_mask.index, unavail_mask.columns].where(~unavail_mask, other=0)

    return contacts, period_dates, availability_all


def import_raw_schedule(folderpath=None, filename='schedule_raw.npy'):
    """
    Import the raw schedule numpy array
    """
    if folderpath is None:
        filepath = os.path.join(parent_fullpath, filename)
    else:
        filepath = os.path.join(folderpath, filename)

    if os.path.exists(filepath):
        print("Importing raw historic schedule")
        bool_schedule = np.load(filepath)
    else:
        print("No historic schedule found")
        bool_schedule = None
    return bool_schedule


def import_schedules(folderpath=None,
                     schedule_filename='schedule.xlsx',
                     raw_schedule_filename='schedule_raw.npy',
                     test=False):
    if folderpath is None:
        schedule_filepath = os.path.join(parent_fullpath, schedule_filename)
        raw_schedule_filepath = os.path.join(parent_fullpath, raw_schedule_filename)
    else:
        schedule_filepath = os.path.join(folderpath, schedule_filename)
        raw_schedule_filepath = os.path.join(folderpath, raw_schedule_filename)

    if test != False:
        print("Deleting previous schedule files")
        if os.path.exists(schedule_filepath):
            os.remove(schedule_filepath)
        if os.path.exists(raw_schedule_filepath):
            os.remove(raw_schedule_filepath)

    if os.path.exists(schedule_filepath):
        print("Existing schedule exists, importing schedule")
        schedule = pd.read_excel(schedule_filepath, index_col=0)
        bool_schedule = generate_boolean_schedule(schedule)
    else:
        print("No existing schedule found. Ensure current_period is set to 1.")
        schedule = None
        bool_schedule = None
    return schedule, bool_schedule


def period_meeting_list(contacts, bool_schedule, period, full=True,
                        save=False, folderpath=None):
    """
    Generate dataframe containing all pairs of people scheduled to meet in a given period.
    """
    save_name = f'period_{period}_meeting_list.xlsx'
    if folderpath is None:
        save_path = os.path.join(parent_fullpath, save_name)
    else:
        save_path = os.path.join(folderpath, save_name)
    
    idxs, jdxs = np.where(bool_schedule[:, :, period - 1] == 1)
    persons1 = contacts.index[idxs]
    persons2 = contacts.index[jdxs]
    if full:
        period_meetings = pd.DataFrame({'Person 1': persons1, 'Person 2': persons2,
                                        'Person 1 email': contacts.loc[persons1, "email"].values,
                                        'Person 2 email': contacts.loc[persons2, "email"].values,
                                        'Person 1 assistant': contacts.loc[persons1, "Assistant"].values,
                                        'Person 2 assistant': contacts.loc[persons2, "Assistant"].values,
                                        'Person 1 assistant email': contacts.loc[persons1, "Assistant email"].values,
                                        'Person 2 assistant email': contacts.loc[persons2, "Assistant email"].values,
                                        'Assistant\'s emails': ["{}; {}".format(a1, a2) for a1, a2 in zip(contacts.loc[persons1, "Assistant email"].values, contacts.loc[persons2, "Assistant email"].values)]
                                        })
        if save:
            period_meetings.to_excel(save_path)
    else:
        period_meetings = pd.DataFrame({'Person 1': persons1, 'Person 2': persons2})

    return period_meetings


def generate_meeting_schedule(contacts, bool_schedule, save=True,
                              folderpath=None, save_name='schedule.xlsx'):
    """
    Generate readable schedule from raw schedule numpy array.
    """
    if folderpath is None:
        save_path = os.path.join(parent_fullpath, save_name)
    else:
        save_path = os.path.join(folderpath, save_name)

    periods = bool_schedule.shape[2]
    schedule = pd.DataFrame("", index=contacts.index, columns=["Period {}".format(x) for x in range(1, periods + 1)])
    for k in range(periods):
        paired_person = period_meeting_person(contacts, bool_schedule, k + 1, contacts.index.values, folderpath=folderpath)
        schedule.loc[:, "Period {}".format(k + 1)] = paired_person
    if save:
        schedule.to_excel(save_path)
    return schedule


def period_meeting_person(contacts, bool_schedule, period, persons, folderpath=None):
    """
    Return the person who is scheduled to meet with a given person in a given period.
    """
    if isinstance(persons, str):
        persons = [persons]
    paired_persons = []
    period_pairs = period_meeting_list(contacts, bool_schedule, period, full=False, folderpath=folderpath)
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


def participant_update(contacts, schedule, bool_schedule):
    if isinstance(schedule, type(None)):
        print("No schedule provided. Skipping participant update")
        return contacts, schedule, bool_schedule

    if (contacts.index.difference(schedule.index).shape[0] == 0) and (schedule.index.difference(contacts.index).shape[0] == 0):
        print("No participant changes")
        return contacts, schedule, bool_schedule
    else:
        for_removal = schedule.index.difference(contacts.index)
        if for_removal.shape[0] > 0:
            print("Removing participants {}".format(for_removal.values))
            ixs = schedule.index.get_indexer(for_removal.values)
            bool_schedule = np.delete(bool_schedule, ixs, axis=0)
            bool_schedule = np.delete(bool_schedule, ixs, axis=1)
            schedule = generate_meeting_schedule(contacts, bool_schedule)
            # schedule = schedule.drop(for_removal, axis=0)
            # schedule = schedule.drop(for_removal, axis=1)

        for_addition = contacts.index.difference(schedule.index)
        if for_addition.shape[0] > 0:
            print("Adding participants {}".format(for_addition.values))
            new_rows = pd.DataFrame(np.nan, index=for_addition.values, columns=schedule.columns)
            schedule = pd.concat([schedule, new_rows])
            # schedule[for_addition.values] = 0
            pad_width = ((0, for_addition.shape[0]), (0, for_addition.shape[0]), (0, 0))
            bool_schedule = np.pad(bool_schedule, pad_width=pad_width, mode='constant', constant_values=0)
        return contacts, schedule, bool_schedule


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


def run_schedule_optimisation(contacts, bool_schedule, n_periods, availability,
                              current_period=1, verbose=False,
                              multiple_meetings='strict'):
    # determine shape of schedule variable
    n_people = contacts.shape[0]
    if isinstance(n_periods, type(None)):
        n_periods = n_people - (n_people % 2 == 0)  # the number periods (e.g. weeks if meetings are weekly) in the season
    total_periods = n_periods + (current_period - 1)

    # Define variable
    X = cp.Variable((n_people, n_people, total_periods), boolean=True)

    # define constraints
    constraints = []

    # Get upper triangle mask (including diagonal)
    upper_mask = np.triu(np.ones((n_people, n_people)), k=0)

    for k in range(total_periods):
        # upper triangle mask will always be 0
        constraints.append(cp.multiply(upper_mask, X[:, :, k]) == 0)

        # each person can only meet once per week, each person is represented by a row and column
        for i in range(n_people):
            constraints.append(cp.sum(X[i, :, k]) + cp.sum(X[:, i, k]) <= 1)

    # set unavailability
    # do this by setting the sum of the corresponding row and column to 0?
    for k in range(total_periods):
        period_unavail_idxs = availability.index.get_indexer(availability[availability[k + 1] == 0].index)
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
        objective = cp.Maximize(cp.sum(X) )
        
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
                        for l in range(k + 1, total_periods):
                            # Enforce Z[i,j,k,l] = X[i,j,k] AND X[i,j,l]
                            constraints += [
                                Z[i, j, k, l] <= X[i, j, k],
                                Z[i, j, k, l] <= X[i, j, l],
                                Z[i, j, k, l] >= X[i, j, k] + X[i, j, l] - 1,
                            ]
        
            penalty_terms = []
        
            for i in range(n_people):
                for j in range(i):
                    for k in range(total_periods):
                        for l in range(k + 1, total_periods):
                            penalty = penalty_weighting(abs(k - l), max_penalty=1, decay_rate=0.1)
                            penalty_terms.append(penalty * Z[i, j, k, l])
        
            ## obective from V0.2.0
            # objective = cp.Maximize(cp.sum(X) - cp.sum(cp.hstack(penalty_terms)))
            
            ## new objective for V0.3.0 that penalises there being no meetings
            objective = cp.Maximize(cp.sum(X) - cp.sum(cp.hstack(penalty_terms)) - total_periods * ((total_periods * n_people**2) - cp.sum(X)))

        else:
            objective = cp.Maximize(cp.sum(X) )
    

    # Solve the problem
    warnings.filterwarnings('ignore')

    problem = cp.Problem(objective, constraints)
    problem.solve(verbose=verbose)

    bool_schedule = (X.value >= 0.5).astype(int)

    # save
    # np.save('/content/drive/MyDrive/CoffeeClub/schedule_raw.npy', bool_schedule)

    return bool_schedule


if __name__ == "__main__":
    # SETUP
    # set below as per usage
    current_period = 1  # generating for this period and future periods. A value greater than 1 will set preceeding periods as fixed e.g. they have already happened.
    n_periods = 7  # set the horizon of the schedule (the term length). If set to None, this will be the minimum number of periods to ensure a meeting between every possible pair is arranged.
    parent_dir = 'example'

    # advanced - change if required
    parent_fullpath = os.path.join('../', parent_dir)

    # create directory if doesn't already exist
    os.makedirs(parent_fullpath, exist_ok=True)

    # import contacts dataframe
    contacts, dates, availability = import_contacts_df(test=4, n_periods=n_periods)

    # import current schedule dataframe
    schedule, bool_schedule = import_schedules(test=True)

    bool_schedule = run_schedule_optimisation(contacts, bool_schedule, n_periods, availability)

    meeting_schedule = generate_meeting_schedule(contacts, bool_schedule, save=True)

    all_meetings_period = period_meeting_list(contacts, bool_schedule, 1, full=False, save=False)
    all_meetings_period
