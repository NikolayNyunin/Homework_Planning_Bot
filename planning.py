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

    with open('files/' + str(user_id) + '_schedule.json', 'w', encoding='utf-8') as schedule_file:
        json.dump(schedule, schedule_file)

    if os.path.exists('files/' + str(user_id) + '_homework.json'):
        os.remove('files/' + str(user_id) + '_homework.json')

    users_path = 'files/users.json'
    if os.path.exists(users_path):
        with open(users_path, encoding='utf-8') as users_json:
            users = json.load(users_json)
    else:
        users = []

    if user_id not in users:
        users.append(user_id)

    with open(users_path, 'w', encoding='utf-8') as users_json:
        json.dump(users, users_json)


def get_schedule(user_id, ordinal_date):
    with open('files/' + str(user_id) + '_subjects.json', encoding='utf-8') as subjects_json:
        subjects = json.load(subjects_json)
    with open('files/' + str(user_id) + '_schedule.json', encoding='utf-8') as schedule_json:
        schedule = json.load(schedule_json)

    try:
        with open('files/' + str(user_id) + '_homework.json', encoding='utf-8') as homework_json:
            homework = json.load(homework_json)
    except FileNotFoundError:
        homework_exists = False
    else:
        homework_exists = True

    date = datetime.date.fromordinal(ordinal_date)
    week = date.isocalendar()[1] % 2
    week_day = date.weekday()
    ordinal_date = str(ordinal_date)

    result = '<i>{} ({}.{}):</i>\n\n'.format(WEEK_DAYS[week_day], str(date.day).zfill(2), str(date.month).zfill(2))
    if week_day == 6:
        result += 'No lessons.\n\n'
    else:
        day_schedule = schedule[week][week_day]
        if day_schedule == [-1] * MAX_LESSONS:
            result += 'No lessons.\n\n'
        else:
            for index in range(MAX_LESSONS):
                subject_index = str(day_schedule[index])

                if subject_index != '-1':
                    result += '{}:   {}\n'.format(index + 1, subjects[subject_index]['Subject'])

                    teacher = subjects[subject_index]['Teacher']
                    if teacher is not None:
                        result += teacher + '\n'

                    room = subjects[subject_index]['Room']
                    if room is not None:
                        result += room + '\n'

                    if homework_exists:
                        if ordinal_date in homework:
                            for h in homework[ordinal_date]:
                                if h['Subject'] == day_schedule[index] and h['For lesson']:
                                    result += '❗<b>{}</b>\n'.format(h['Description'])

                    result += '\n'

    if homework_exists:
        if ordinal_date in homework:
            for h in homework[ordinal_date]:
                if not h['For lesson']:
                    result += subjects[str(h['Subject'])]['Subject'] + '\n'
                    result += '❗<b>{}</b>\n\n'.format(h['Description'])

    return result


def in_schedule(user_id, ordinal_date, subject_index):
    with open('files/' + str(user_id) + '_schedule.json', encoding='utf-8') as schedule_json:
        schedule = json.load(schedule_json)

    date = datetime.date.fromordinal(ordinal_date)
    week = date.isocalendar()[1] % 2
    week_day = date.weekday()
    if week_day == 6:
        return False

    return subject_index in schedule[week][week_day]


def get_subjects(user_id):
    with open('files/' + str(user_id) + '_subjects.json', encoding='utf-8') as subjects_json:
        subjects = json.load(subjects_json)
    subjects = [val['Subject'] for val in subjects.values()]
    return subjects


def add_homework(user_id, subject_index, date, for_lesson, description):
    path = 'files/' + str(user_id) + '_homework.json'
    if os.path.exists(path):
        with open(path, encoding='utf-8') as homework_json:
            homework = json.load(homework_json)
    else:
        homework = {}

    with open('files/' + str(user_id) + '_schedule.json', encoding='utf-8') as schedule_json:
        schedule = json.load(schedule_json)

    if not date:
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
                str_date = str(date)
                if str_date not in homework:
                    homework[str_date] = []
                homework[str_date].append({'Subject': subject_index, 'For lesson': for_lesson,
                                           'Description': description})
                break

            date += 1
            week_day += 1

    else:
        date = str(date)
        if date not in homework:
            homework[date] = []
        homework[date].append({'Subject': subject_index, 'For lesson': for_lesson, 'Description': description})

    with open(path, 'w', encoding='utf-8') as homework_json:
        json.dump(homework, homework_json, ensure_ascii=False, indent=4)


def delete_past_homework():
    date = datetime.datetime.now(TIMEZONE).date().toordinal()

    users_path = 'files/users.json'
    if os.path.exists(users_path):
        with open(users_path, encoding='utf-8') as users_json:
            users = json.load(users_json)

        for user in users:
            homework_path = 'files/{}_homework.json'.format(user)
            if os.path.exists(homework_path):
                with open(homework_path, encoding='utf-8') as homework_json:
                    homework = json.load(homework_json)

                for d in sorted(map(int, homework.keys())):
                    if d < date - 1:
                        del homework[d]
                    else:
                        break

                with open(homework_path, 'w', encoding='utf-8') as homework_json:
                    json.dump(homework, homework_json, ensure_ascii=False, indent=4)
