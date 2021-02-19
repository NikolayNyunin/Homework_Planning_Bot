import pandas as pd
import json
from itertools import islice

MAX_LESSONS = 5


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


def get_schedule(user_id, week, day):
    with open('files/' + str(user_id) + '_subjects.json', 'r', encoding='utf-8') as subjects_json:
        subjects = json.load(subjects_json)
    with open('files/' + str(user_id) + '_schedule.json', 'r') as schedule_json:
        schedule = json.load(schedule_json)

    result = ''
    lessons = schedule[week][day]
    for index in range(MAX_LESSONS):
        result += str(index + 1) + ':\t'
        if lessons[index] == -1:
            result += 'No lesson.\n\n'
        else:
            result += subjects[str(lessons[index])]['Subject'] + '\n\n'

    return result
