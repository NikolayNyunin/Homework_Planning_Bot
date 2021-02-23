import os
import datetime
import pytz
import json
from itertools import islice

import pandas as pd

MAX_LESSONS = 5
WEEK_DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
TIMEZONE = pytz.timezone('Europe/Moscow')


def set_schedule(user_id, file):
    file = pd.ExcelFile(file)
    df = file.parse(file.sheet_names[0])

    subjects = df[df.columns[:3]].to_dict('index')

    for key1 in subjects.keys():
        for key2, val in subjects[key1].items():
            if type(val) == float:
                subjects[key1][key2] = None

    with open('files/' + str(user_id) + '_subjects.json', 'w', encoding='utf-8') as subjects_file:
        json.dump(subjects, subjects_file, ensure_ascii=False, indent=4)

    schedule = [[[], [], [], [], [], []],
                [[], [], [], [], [], []]]

    day_index = 0
    for label, values in islice(df.items(), 5, 11):
        for subject in values[1:]:
            if type(subject) == float:
                schedule[0][day_index].append(-1)
            else:
                schedule[0][day_index].append(subject)
            if len(schedule[0][day_index]) >= MAX_LESSONS:
                break
        day_index += 1

    day_index = 0
    for label, values in islice(df.items(), 13, 19):
        for subject in values[1:]:
            if type(subject) == float:
                schedule[1][day_index].append(-1)
            else:
                schedule[1][day_index].append(subject)
            if len(schedule[1][day_index]) >= MAX_LESSONS:
                break
        day_index += 1

    with open('files/' + str(user_id) + '_schedule.json', 'w') as schedule_file:
        json.dump(schedule, schedule_file)

    if os.path.exists('files/' + str(user_id) + '_homework.json'):
        os.remove('files/' + str(user_id) + '_homework.json')


def get_schedule(user_id, ordinal_date):
    with open('files/' + str(user_id) + '_subjects.json', 'r', encoding='utf-8') as subjects_json:
        subjects = json.load(subjects_json)
    with open('files/' + str(user_id) + '_schedule.json', 'r') as schedule_json:
        schedule = json.load(schedule_json)

    try:
        with open('files/' + str(user_id) + '_homework.json', 'r', encoding='utf-8') as homework_json:
            homework = json.load(homework_json)
    except FileNotFoundError:
        homework_exists = False
    else:
        homework_exists = True

    date = datetime.date.fromordinal(ordinal_date)
    week = date.isocalendar()[1] % 2
    week_day = date.weekday()

    result = '{0} ({1}.{2}):\n\n'.format(WEEK_DAYS[week_day], date.day, str(date.month).zfill(2))
    if week_day == 6:
        return result + 'No lessons.'

    day_schedule = schedule[week][week_day]
    for index in range(MAX_LESSONS):
        result += str(index + 1) + ':\t'
        if day_schedule[index] == -1:
            result += 'No lesson.\n\n'
        else:
            result += subjects[str(day_schedule[index])]['Subject'] + '\n'
            if homework_exists:
                for h in homework:
                    if h['Date'] == ordinal_date and h['Subject'] == day_schedule[index]:
                        result += h['Description'] + '\n'

            result += '\n'

    return result


def get_subjects(user_id):
    with open('files/' + str(user_id) + '_subjects.json', 'r', encoding='utf-8') as subjects_json:
        subjects = json.load(subjects_json)
    subjects = [val['Subject'] for val in subjects.values()]
    return subjects


def add_homework(user_id, subject_index, description):
    path = 'files/' + str(user_id) + '_homework.json'
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as homework_json:
            homework = json.load(homework_json)
    else:
        homework = []

    with open('files/' + str(user_id) + '_schedule.json', 'r') as schedule_json:
        schedule = json.load(schedule_json)

    date = datetime.datetime.now(TIMEZONE).date().toordinal() + 1
    week = datetime.date.fromordinal(date).isocalendar()[1] % 2
    week_day = datetime.date.fromordinal(date).weekday()
    for i in range(14):
        if week_day == 6:
            date += 1
            week = (week + 1) % 2
            week_day = 0
            continue

        day_schedule = schedule[week][week_day]
        if subject_index in day_schedule:
            homework.append({'Date': date, 'Subject': subject_index, 'Description': description})
            break

        date += 1
        week_day += 1

    with open(path, 'w', encoding='utf-8') as homework_json:
        json.dump(homework, homework_json, ensure_ascii=False)
