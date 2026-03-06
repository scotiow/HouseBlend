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

Recommended API (class-first):
    parliament = Parliament.from_excel(parliament_name="example")
    parliament.update_participants().optimise(n_to_schedule=5)
    parliament.build_schedule(save=True)
    participants_df, assistants_df = parliament.export_mailmerge(period=1, save=True)

Migration note:
    Function-style APIs are retained for transition convenience, but new code should
    use Parliament methods as the primary interface.

Copyright (c) 2025 Scot Wheeler

This file is part of HouseBlend, which is licensed under the MIT License.
You may obtain a copy of the License at
https://opensource.org/licenses/MIT

"""

__author__ = "Scot Wheeler"
__license__ = "MIT"
__version__ = "0.6.0"

import numpy as np
import cvxpy as cp
import pandas as pd
import datetime as dt
import os
from faker import Faker
import logging
from dataclasses import dataclass
from typing import Optional, Tuple
logger = logging.getLogger(__name__)


@dataclass
class ParliamentState:
    """Container for the in-memory scheduling state.

    Invariants are checked via ``validate()`` and are designed around the Excel-based
    Hansard model:
    - ``contacts`` index is the canonical participant list.
    - ``availability`` and ``schedule`` have matching participant index and period columns.
    - ``bool_schedule`` has shape (n_people, n_people, n_periods).
    """

    contacts: pd.DataFrame
    dates: pd.DataFrame
    availability: pd.DataFrame
    schedule: pd.DataFrame
    bool_schedule: np.ndarray

    def as_tuple(self) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, np.ndarray]:
        return self.contacts, self.dates, self.availability, self.schedule, self.bool_schedule

    def validate(self) -> None:
        """Validate consistency between tabular state and bool schedule."""
        n_people = self.contacts.shape[0]
        n_periods = self.availability.shape[1]

        if self.bool_schedule.ndim != 3:
            raise ValueError("bool_schedule must be a 3D array")

        if self.bool_schedule.shape[0] != n_people or self.bool_schedule.shape[1] != n_people:
            raise ValueError("bool_schedule participant dimensions must match contacts rows")

        if self.bool_schedule.shape[2] != n_periods:
            raise ValueError("bool_schedule periods must match availability columns")

        if not self.schedule.index.equals(self.contacts.index):
            raise ValueError("schedule index must match contacts index")

        if not self.availability.index.equals(self.contacts.index):
            raise ValueError("availability index must match contacts index")

        if self.schedule.shape[1] != n_periods:
            raise ValueError("schedule period columns must match availability period columns")

        expected_period_cols = [f"Period {i}" for i in range(1, n_periods + 1)]
        if list(self.availability.columns) != expected_period_cols:
            raise ValueError("availability columns must follow exact 'Period X' sequence")

        if list(self.schedule.columns) != expected_period_cols:
            raise ValueError("schedule columns must follow exact 'Period X' sequence")

        if not np.isin(self.bool_schedule, [0, 1]).all():
            raise ValueError("bool_schedule must contain only 0/1 values")

        for k in range(self.bool_schedule.shape[2]):
            layer = self.bool_schedule[:, :, k]
            if not np.all(np.diag(layer) == 0):
                raise ValueError("bool_schedule diagonal must be zero for all periods")
            if not np.all(np.triu(layer, k=0) == 0):
                raise ValueError("bool_schedule must only use strict lower triangle entries")

        if not self.dates.index.is_unique:
            raise ValueError("dates index must be unique")


class HansardRepository:
    """Persistence service for Hansard workbook I/O."""

    def _resolve_paths(self, folderpath, filename, parliament_name):
        if folderpath is None:
            folderpath = parliament_name
        if filename is None:
            filename = f'{parliament_name}_hansard.xlsx'
        if not filename.endswith('.xlsx'):
            filename += '.xlsx'
        return folderpath, filename

    def load(self, folderpath=None, filename=None, parliament_name="example", test=False, n_periods=None):
        if test is not False:
            return self.create_initial(
                n_participants=test if isinstance(test, int) else 4,
                parliament_name=parliament_name,
                folderpath=folderpath,
                filename=filename,
                n_periods=n_periods,
            )

        folderpath, filename = self._resolve_paths(folderpath, filename, parliament_name)
        filepath = os.path.join(folderpath, filename)

        if os.path.exists(filepath):
            logger.info("Importing contacts")
            contacts = pd.read_excel(filepath, index_col=0, sheet_name="Participants")
            dates = pd.read_excel(filepath, index_col=0, sheet_name="Dates")
            availability = pd.read_excel(filepath, index_col=0, sheet_name="Availability")
            availability = availability.fillna(1)
            schedule = pd.read_excel(filepath, index_col=0, sheet_name="Schedule")
            bool_schedule = self._generate_boolean_schedule(schedule)
            return contacts, dates, availability, schedule, bool_schedule
        else:
            print(f"No hansard file found at {filepath}")
            return None

    def save(self, contacts, dates, availability, schedule, parliament_name="example", folderpath=None, filename=None):
        folderpath, filename = self._resolve_paths(folderpath, filename, parliament_name)
        os.makedirs(folderpath, exist_ok=True)
        filepath = os.path.join(folderpath, filename)

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

    def _generate_test_contacts(self, n_participants):
        participant_names = []
        fake = Faker()
        for _ in range(n_participants):
            participant_names.append(fake.first_name() + " " + fake.last_name())
        return pd.DataFrame({
            "Person": participant_names,
            "email": [f"{name.replace(' ', '')}@email.com" for name in participant_names],
            "Assistant": participant_names,
            "Assistant email": [f"{name.replace(' ', '')}@email.com" for name in participant_names],
        }).set_index("Person")

    def _generate_boolean_schedule(self, schedule):
        n_people = int(schedule.shape[0])
        n_periods = int(schedule.shape[1])
        bool_schedule = np.zeros((n_people, n_people, n_periods))
        for k in range(n_periods):
            idxs = schedule.index.get_indexer(schedule[f"Period {k + 1}"].values)
            jdxs = np.where(idxs != -1)[0]
            bool_schedule[idxs[idxs != -1], jdxs, k] = 1
            lower_mask = np.tril(np.ones((n_people, n_people)), k=0)
            bool_schedule[:, :, k] = bool_schedule[:, :, k] * lower_mask
        return bool_schedule

    def _min_periods(self, n_people):
        return n_people - (n_people % 2 == 0)

    def create_initial(self, n_participants, parliament_name="example", folderpath=None, filename=None, n_periods=None):
        folderpath, filename = self._resolve_paths(folderpath, filename, parliament_name)
        os.makedirs(folderpath, exist_ok=True)
        filepath = os.path.abspath(os.path.join(folderpath, filename))

        if os.path.exists(filepath):
            print("Deleting previous hansard file")
            os.remove(filepath)

        print(f"Creating new parliament of {n_participants} participants at {filepath}")
        contacts = self._generate_test_contacts(n_participants)

        if n_periods is None:
            n_periods = self._min_periods(n_participants)

        start_date = (dt.datetime.today() + dt.timedelta(days=(7 - dt.datetime.today().weekday()))).date()
        period_dates = pd.DataFrame({
            "Start Date": pd.date_range(start_date, periods=n_periods, freq='2W').values,
            "End Date": pd.date_range(start_date + dt.timedelta(days=14), periods=n_periods, freq='2W').values
        }, index=list(range(1, n_periods + 1)))
        period_dates.index.name = "Period"

        availability = pd.DataFrame(np.ones((contacts.shape[0], n_periods)), 
                                    columns=[f"Period {i}" for i in range(1, n_periods + 1)], 
                                    index=contacts.index)
        schedule = pd.DataFrame("", index=contacts.index, 
                               columns=[f"Period {x}" for x in range(1, n_periods + 1)])

        self.save(contacts, period_dates, availability, schedule, 
                  parliament_name=parliament_name, folderpath=folderpath, filename=filename)

        return contacts, period_dates, availability, schedule, self._generate_boolean_schedule(schedule)


class ScheduleBuilder:
    """Service for schedule transformations and period views."""

    def _generate_boolean_schedule(self, schedule):
        n_people = int(schedule.shape[0])
        n_periods = int(schedule.shape[1])
        bool_schedule = np.zeros((n_people, n_people, n_periods))
        for k in range(n_periods):
            idxs = schedule.index.get_indexer(schedule[f"Period {k + 1}"].values)
            jdxs = np.where(idxs != -1)[0]
            bool_schedule[idxs[idxs != -1], jdxs, k] = 1
            lower_mask = np.tril(np.ones((n_people, n_people)), k=0)
            bool_schedule[:, :, k] = bool_schedule[:, :, k] * lower_mask
        return bool_schedule

    def period_meeting_list(self, contacts, bool_schedule, period, save=False, folderpath=None, parliament_name='example'):
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

    def period_meeting_person(self, contacts, bool_schedule, period, persons):
        if isinstance(persons, str):
            persons = [persons]
        paired_persons = []
        period_pairs = self.period_meeting_list(contacts, bool_schedule, period)
        for person in persons:
            mask = period_pairs == person
            if mask.sum().sum() == 0:
                paired_persons.append(np.nan)
            else:
                pair = period_pairs.loc[mask.sum(axis=1) == 1, :].to_numpy().flatten()
                paired_persons.append(pair[pair != person][0])
        return paired_persons if len(paired_persons) > 1 else paired_persons[0]

    def generate_meeting_schedule(self, contacts, dates, availability, bool_schedule, save=False,
                                  folderpath=None, parliament_name="example", filename=None):
        periods = bool_schedule.shape[2]
        schedule = pd.DataFrame("", index=contacts.index, columns=[f"Period {x}" for x in range(1, periods + 1)])
        for k in range(periods):
            paired_person = self.period_meeting_person(contacts, bool_schedule, k + 1, contacts.index.values)
            schedule.loc[:, f"Period {k + 1}"] = paired_person
        
        if save:
            repo = HansardRepository()
            repo.save(contacts, dates, availability, schedule, 
                     parliament_name=parliament_name, folderpath=folderpath, filename=filename)
        return schedule


class SchedulerOptimizer:
    """Service for optimisation and date/schedule horizon maintenance."""

    def _penalty_weighting(self, difference, max_penalty, decay_rate=1):
        return max_penalty * np.exp(-decay_rate * difference)

    def _min_periods(self, n_people):
        return n_people - (n_people % 2 == 0)

    def check_dates_and_add(self, current_period, n_to_schedule, dates, availability, schedule,
                            bool_schedule, parliament_name, folderpath=None, filename=None):
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
                last_start_date = dates["Start Date"].max()
                last_end_date = dates["End Date"].max()
                new_start_date = last_start_date + pd.tseries.frequencies.to_offset(freq)
                new_end_date = last_end_date + pd.tseries.frequencies.to_offset(freq)
                new_dates = pd.DataFrame({"Start Date": [new_start_date], "End Date": [new_end_date]}, index=[i])
                dates = pd.concat([dates, new_dates])
        
        for i in range(current_period, current_period + n_to_schedule):
            if f"Period {i}" not in availability.columns:
                availability[f"Period {i}"] = 1
            if f"Period {i}" not in schedule.columns:
                schedule[f"Period {i}"] = ""
        
        if bool_schedule.shape[2] < current_period + n_to_schedule - 1:
            additional_layers = current_period + n_to_schedule - 1 - bool_schedule.shape[2]
            bool_schedule = np.concatenate([bool_schedule, np.zeros((bool_schedule.shape[0], bool_schedule.shape[1], additional_layers))], axis=2)

        return dates, availability, schedule, bool_schedule

    def run_optimisation(self, contacts, dates, availability, schedule, bool_schedule, n_to_schedule,
                         current_period=1, verbose=False, multiple_meetings='strict', save=False,
                         folderpath=None, parliament_name="example", filename=None, iterative_limit=3):
        n_people = contacts.shape[0]

        if isinstance(n_to_schedule, type(None)):
            n_to_schedule = self._min_periods(n_people=n_people)

        dates, availability, schedule, bool_schedule = self.check_dates_and_add(
            current_period, n_to_schedule, dates, availability, schedule, bool_schedule, 
            parliament_name, folderpath=folderpath, filename=filename
        )

        counter = 0
        if n_to_schedule > iterative_limit:
            print(f"Number of periods to schedule ({n_to_schedule}) is greater than iterative limit ({iterative_limit}). Running optimisation iteratively in chunks of {iterative_limit} periods.")
        
        while counter < (n_to_schedule // iterative_limit) + 1:
            n_it = min(iterative_limit, n_to_schedule - counter * iterative_limit)
            start_period = current_period + counter * iterative_limit
            
            print(f"Running optimisation for periods {start_period} to {start_period + n_it - 1}")

            X = cp.Variable((n_people, n_people, n_it), boolean=True)
            constraints = []
            upper_mask = np.triu(np.ones((n_people, n_people)), k=0)

            for k in range(n_it):
                constraints.append(cp.multiply(upper_mask, X[:, :, k]) == 0)
                for i in range(n_people):
                    constraints.append(cp.sum(X[i, :, k]) + cp.sum(X[:, i, k]) <= 1)

            for k, m in zip(range(n_it), range(start_period - 1, start_period - 1 + n_it)):
                period_unavail_idxs = availability.index.get_indexer(availability[availability[f'Period {m + 1}'] == 0].index)
                constraints.append(cp.sum(X[period_unavail_idxs, :, k]) == 0)
                constraints.append(cp.sum(X[:, period_unavail_idxs, k]) == 0)

            weighting = np.zeros((n_people, n_people))
            for i in range(n_people):
                for j in range(i):
                    if i != j:
                        last_meeting_period = 0
                        for k in range(start_period - 1, -1, -1):
                            if (bool_schedule[i, j, k] == 1) or (bool_schedule[j, i, k] == 1):
                                last_meeting_period = max(last_meeting_period, k + 1)
                        weighting[i, j] = 1/self._penalty_weighting(abs(start_period - last_meeting_period), max_penalty=1, decay_rate=0.1)

            if n_it > 1:
                Z = cp.Variable((n_people, n_people, n_it, n_it), boolean=True)

                for i in range(n_people):
                    for j in range(i):
                        for k in range(n_it):
                            for L in range(k + 1, n_it):
                                constraints += [
                                    Z[i, j, k, L] <= X[i, j, k],
                                    Z[i, j, k, L] <= X[i, j, L],
                                    Z[i, j, k, L] >= X[i, j, k] + X[i, j, L] - 1,
                                ]

                penalty_terms = []
                for i in range(n_people):
                    for j in range(i):
                        for k in range(n_it):
                            for L in range(k + 1, n_it):
                                penalty = self._penalty_weighting(abs(k - L), max_penalty=1, decay_rate=0.1)
                                penalty_terms.append(penalty * Z[i, j, k, L])

                weighting_expanded = np.zeros((n_people, n_people, n_it))
                for k in range(n_it):
                    weighting_expanded[:, :, k] = weighting
                objective = cp.Maximize(cp.sum(cp.multiply(weighting_expanded, X[:, :, :])) - cp.sum(penalty_terms))
            else:
                weighting_expanded = np.zeros((n_people, n_people, n_it))
                for k in range(n_it):
                    weighting_expanded[:, :, k] = weighting
                print("Single period, only including past meeting penalty")
                objective = cp.Maximize(cp.sum(cp.multiply(weighting_expanded, X[:, :, :])))

            problem = cp.Problem(objective, constraints)
            problem.solve(verbose=verbose)

            new_schedule = (X.value >= 0.5).astype(int)
            bool_schedule[:, :, start_period - 1:start_period - 1 + n_it] = new_schedule
            
            counter += 1

        if save:
            builder = ScheduleBuilder()
            schedule = builder.generate_meeting_schedule(
                contacts, dates, availability, bool_schedule, parliament_name=parliament_name, 
                folderpath=folderpath, filename=filename
            )
            repo = HansardRepository()
            repo.save(contacts, dates, availability, schedule, 
                     parliament_name=parliament_name, folderpath=folderpath, filename=filename)

        return bool_schedule


class MailmergeExporter:
    """Service for mailmerge exports and helper formatting."""

    def _create_mailto_link(self, meeting, subject=None):
        to = f"{meeting['Person 1 assistant email']};{meeting['Person 2 assistant email']}"
        cc = f"{meeting['Person 1 email']};{meeting['Person 2 email']}"
        link = f"mailto:{to}?cc={cc}"
        if subject:
            link += f"&subject={subject.replace(' ', '%20')}"
        return link

    def _get_first_name(self, name):
        if isinstance(name, str) and name.strip():
            return name.strip().split()[0]
        return ""

    def _format_date_with_suffix(self, date):
        if pd.isnull(date):
            return None
        date = pd.to_datetime(date)
        day = date.day
        if 10 <= day % 100 <= 20:
            suffix = 'th'
        else:
            suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
        return date.strftime(f"%A {day}{suffix} %B %Y")

    def export_for_mailmerge(self, contacts, bool_schedule, period, dates,
                             save=False, folderpath=None, parliament_name='example',
                             save_prefix="HouseBlend", subject=None):
        idxs, jdxs = np.where(bool_schedule[:, :, period - 1] == 1)
        persons1 = contacts.index[idxs]
        persons2 = contacts.index[jdxs]
        
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
            'Start Date Long': self._format_date_with_suffix(dates.loc[period, "Start Date"]),
            'End Date Long': self._format_date_with_suffix(dates.loc[period, "End Date"])
        })
        
        participants_rows = []
        for _, meeting in meetings_base.iterrows():
            participants_rows.append({
                'mailto full name': meeting['Person 1'],
                'mailto first name': self._get_first_name(meeting['Person 1']),
                'mailto assistant full name': meeting['Person 1 assistant'],
                'paired with full name': meeting['Person 2'],
                'paired with first name': self._get_first_name(meeting['Person 2']),
                'start date': meeting['Start Date'],
                'end date': meeting['End Date'],
                'start date long': meeting['Start Date Long'],
                'end date long': meeting['End Date Long'],
                'mailto email': meeting['Person 1 email']
            })
            participants_rows.append({
                'mailto full name': meeting['Person 2'],
                'mailto first name': self._get_first_name(meeting['Person 2']),
                'mailto assistant full name': meeting['Person 2 assistant'],
                'paired with full name': meeting['Person 1'],
                'paired with first name': self._get_first_name(meeting['Person 1']),
                'start date': meeting['Start Date'],
                'end date': meeting['End Date'],
                'start date long': meeting['Start Date Long'],
                'end date long': meeting['End Date Long'],
                'mailto email': meeting['Person 2 email']
            })
        
        participants_df = pd.DataFrame(participants_rows)
        
        assistants_rows = []
        for _, meeting in meetings_base.iterrows():
            assistants_rows.append({
                'mailto assistant full name': meeting['Person 1 assistant'],
                'mailto assistant first name': self._get_first_name(meeting['Person 1 assistant']),
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
            assistants_rows.append({
                'mailto assistant full name': meeting['Person 2 assistant'],
                'mailto assistant first name': self._get_first_name(meeting['Person 2 assistant']),
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
        
        if save:
            save_name = f'{parliament_name}_{save_prefix}_period_{period}_mailmerge.xlsx'
            if folderpath is None:
                folderpath = parliament_name
            save_path = os.path.join(folderpath, save_name)
            with pd.ExcelWriter(save_path) as writer:
                participants_df.to_excel(writer, sheet_name='Participants', index=False)
                assistants_df.to_excel(writer, sheet_name='Assistants', index=False)
        
        return participants_df, assistants_df


class HouseBlendSession:
    """Object-oriented interface for the HouseBlend scheduling workflow.

    Preferred usage is class-first and Excel-first:

    - ``HouseBlendSession.from_excel(...)`` to load an existing Hansard workbook.
    - ``HouseBlendSession.create_excel(...)`` to create a new workbook.
    - ``optimise(...)``, ``build_schedule(...)``, ``export_mailmerge(...)`` for workflow steps.
    """

    def __init__(self, parliament_name="example", folderpath=None, filename=None, state: Optional[ParliamentState] = None,
                 repository: Optional[HansardRepository] = None,
                 scheduler_optimizer: Optional[SchedulerOptimizer] = None,
                 schedule_builder: Optional[ScheduleBuilder] = None,
                 mailmerge_exporter: Optional[MailmergeExporter] = None):
        self.parliament_name = parliament_name
        self.folderpath = folderpath
        self.filename = filename
        self.state = state
        self.repository = repository if repository is not None else HansardRepository()
        self.scheduler_optimizer = scheduler_optimizer if scheduler_optimizer is not None else SchedulerOptimizer()
        self.schedule_builder = schedule_builder if schedule_builder is not None else ScheduleBuilder()
        self.mailmerge_exporter = mailmerge_exporter if mailmerge_exporter is not None else MailmergeExporter()

    @classmethod
    def from_excel(cls, folderpath=None, filename=None, parliament_name="example", test=False, n_periods=None):
        session = cls(parliament_name=parliament_name, folderpath=folderpath, filename=filename)
        session.load(test=test, n_periods=n_periods)
        return session

    @classmethod
    def create_excel(cls, n_participants, parliament_name="example", folderpath=None, filename=None, n_periods=None):
        session = cls(parliament_name=parliament_name, folderpath=folderpath, filename=filename)
        state_tuple = session.repository.create_initial(
            n_participants=n_participants,
            parliament_name=parliament_name,
            folderpath=folderpath,
            filename=filename,
            n_periods=n_periods,
        )
        return session._set_state(state_tuple, validate=False)
    @classmethod
    def from_hansard(cls, *args, **kwargs):
        """Deprecated alias for from_excel()."""
        return cls.from_excel(*args, **kwargs)

    @classmethod
    def create_new(cls, *args, **kwargs):
        """Deprecated alias for create_excel()."""
        return cls.create_excel(*args, **kwargs)
    def _set_state(self, state_tuple, validate=False):
        self.state = ParliamentState(*state_tuple)
        if validate:
            self.state.validate()
        return self

    def load(self, test=False, n_periods=None):
        state_tuple = self.repository.load(
            folderpath=self.folderpath,
            filename=self.filename,
            parliament_name=self.parliament_name,
            test=test,
            n_periods=n_periods,
        )
        if state_tuple is None:
            raise FileNotFoundError("No hansard file found for the configured parliament")
        return self._set_state(state_tuple, validate=False)

    def _require_state(self) -> ParliamentState:
        if self.state is None:
            raise ValueError("No state loaded. Use load(), from_hansard(), or create_new() first.")
        return self.state

    def validate(self):
        self._require_state().validate()
        return self

    def as_tuple(self):
        return self._require_state().as_tuple()

    @property
    def contacts(self):
        return self._require_state().contacts

    @property
    def dates(self):
        return self._require_state().dates

    @property
    def availability(self):
        return self._require_state().availability

    @property
    def schedule(self):
        return self._require_state().schedule

    @property
    def bool_schedule(self):
        return self._require_state().bool_schedule

    def save_excel(self):
        state = self._require_state()
        self.repository.save(
            state.contacts,
            state.dates,
            state.availability,
            state.schedule,
            parliament_name=self.parliament_name,
            folderpath=self.folderpath,
            filename=self.filename,
        )
        return self

    def save(self):
        """Deprecated alias for save_excel()."""
        return self.save_excel()

    def update_participants(self):
        state = self._require_state()
        
        if state.schedule is None:
            print("No schedule provided. Skipping participant update")
            return self

        if (state.contacts.index.difference(state.schedule.index).shape[0] == 0) and \
           (state.schedule.index.difference(state.contacts.index).shape[0] == 0):
            print("No participant changes")
            return self
        
        for_removal = state.schedule.index.difference(state.contacts.index)
        if for_removal.shape[0] > 0:
            print(f"Removing participants {for_removal.values}")
            ixs = state.schedule.index.get_indexer(for_removal.values)
            state.bool_schedule = np.delete(state.bool_schedule, ixs, axis=0)
            state.bool_schedule = np.delete(state.bool_schedule, ixs, axis=1)

        for_addition = state.contacts.index.difference(state.schedule.index)
        if for_addition.shape[0] > 0:
            print(f"Adding participants {for_addition.values}")
            for participant in for_addition:
                contacts_idx = state.contacts.index.get_indexer([participant])
                state.bool_schedule = np.insert(state.bool_schedule, contacts_idx[0], 0, axis=0)
                state.bool_schedule = np.insert(state.bool_schedule, contacts_idx[0], 0, axis=1)

        state.availability = state.availability.loc[state.availability.index.intersection(state.contacts.index), :]
        
        builder = ScheduleBuilder()
        state.schedule = builder.generate_meeting_schedule(
            state.contacts, state.dates, state.availability, state.bool_schedule, 
            save=False, parliament_name=self.parliament_name, folderpath=self.folderpath, filename=self.filename
        )
        
        for participant in for_addition:
            state.availability.loc[participant, :] = 1

        return self

    def ensure_dates(self, current_period, n_to_schedule):
        state = self._require_state()
        dates, availability, schedule, bool_schedule = self.scheduler_optimizer.check_dates_and_add(
            current_period,
            n_to_schedule,
            state.dates,
            state.availability,
            state.schedule,
            state.bool_schedule,
            parliament_name=self.parliament_name,
            folderpath=self.folderpath,
            filename=self.filename,
        )
        return self._set_state((state.contacts, dates, availability, schedule, bool_schedule), validate=False)

    def optimise(self, n_to_schedule, current_period=1, verbose=False,
                 multiple_meetings='strict', save=False, iterative_limit=3):
        state = self._require_state()
        state.bool_schedule = self.scheduler_optimizer.run_optimisation(
            state.contacts,
            state.dates,
            state.availability,
            state.schedule,
            state.bool_schedule,
            n_to_schedule,
            current_period=current_period,
            verbose=verbose,
            multiple_meetings=multiple_meetings,
            save=save,
            folderpath=self.folderpath,
            parliament_name=self.parliament_name,
            filename=self.filename,
            iterative_limit=iterative_limit,
        )
        if save:
            state.schedule = self.schedule_builder.generate_meeting_schedule(
                state.contacts,
                state.dates,
                state.availability,
                state.bool_schedule,
                save=False,
            )
        return self

    def optimize(self, *args, **kwargs):
        """US spelling alias for optimise."""
        return self.optimise(*args, **kwargs)

    def build_schedule(self, save=False):
        state = self._require_state()
        state.schedule = self.schedule_builder.generate_meeting_schedule(
            state.contacts,
            state.dates,
            state.availability,
            state.bool_schedule,
            save=save,
            folderpath=self.folderpath,
            parliament_name=self.parliament_name,
            filename=self.filename,
        )
        return state.schedule

    def export_mailmerge(self, period, save=False, save_prefix="HouseBlend", subject=None):
        state = self._require_state()
        return self.mailmerge_exporter.export_for_mailmerge(
            state.contacts,
            state.bool_schedule,
            period,
            state.dates,
            save=save,
            folderpath=self.folderpath,
            parliament_name=self.parliament_name,
            save_prefix=save_prefix,
            subject=subject,
        )


if __name__ == "__main__":
    # SETUP
    n_to_schedule = 5  # total number of periods to schedule. Set to None for minimum number to satisfy all possible meetings.
    current_period = 1  # the period from which to schedule. Set to 1 for first usage.
    parliament_name = 'example'  # name of parliament - used to create folder for input/output files

    session = HouseBlendSession.from_excel(parliament_name=parliament_name, test=6)

    session.update_participants().optimise(
        n_to_schedule=n_to_schedule,
        current_period=current_period,
        save=True,
    )
    session.build_schedule(save=True)

    session.export_mailmerge(
        period=1,
        save=True,
        save_prefix='HouseBlend',
        subject=None,
    )
