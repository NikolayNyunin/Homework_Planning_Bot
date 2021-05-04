import datetime
import json
from itertools import islice

import pytz
from pandas import ExcelFile

from db import MAX_LESSONS, User, Subject, Homework, Session, get_user

WEEK_DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
SHORT_WEEK_DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
TIMEZONE = pytz.timezone('Europe/Moscow')


class ScheduleNotFoundError(Exception):
    def __str__(self):
        return 'Schedule not found.\nPlease set your schedule before requesting it'


def set_schedule(user_id, file):
    file = ExcelFile(file)
    df = file.parse(file.sheet_names[0])

    data = df[df.columns[1:4]].to_dict('records')

    session, user = get_user(user_id)

    user.subjects.clear()

    for row in data:
        if type(row['Subject']) != float:
            teacher = row['Teacher'] if type(row['Teacher']) != float else None
            room = row['Room'] if type(row['Room']) != float else None
            user.subjects.append(Subject(name=row['Subject'], teacher=teacher, room=room))

    schedule = [[[], [], [], [], [], []],
                [[], [], [], [], [], []]]

    day_index = 0
    for label, values in islice(df.items(), 6, 12):
        for subject in values[1:]:
            if type(subject) == float:
                schedule[0][day_index].append(-1)
            else:
                schedule[0][day_index].append(subject)
            if len(schedule[0][day_index]) >= MAX_LESSONS:
                break
        day_index += 1

    day_index = 0
    for label, values in islice(df.items(), 14, 20):
        for subject in values[1:]:
            if type(subject) == float:
                schedule[1][day_index].append(-1)
            else:
                schedule[1][day_index].append(subject)
            if len(schedule[1][day_index]) >= MAX_LESSONS:
                break
        day_index += 1

    user.schedule = json.dumps(schedule)
    user.homework.clear()

    session.commit()
    session.close()


def get_schedule(user_id, ordinal_date):
    session, user = get_user(user_id)
    if not user.schedule or not user.subjects:
        session.close()
        raise ScheduleNotFoundError

    schedule = json.loads(user.schedule)

    date = datetime.date.fromordinal(ordinal_date)
    week = date.isocalendar()[1] % 2
    week_day = date.weekday()

    result = '<i>{} ({}.{}):</i>\n\n'.format(WEEK_DAYS[week_day], str(date.day).zfill(2), str(date.month).zfill(2))
    if week_day == 6:
        result += 'No lessons.\n\n'
    else:
        day_schedule = schedule[week][week_day]
        if day_schedule == [-1] * MAX_LESSONS:
            result += 'No lessons.\n\n'
        else:
            for index in range(MAX_LESSONS):
                subject_index = day_schedule[index]

                if subject_index != -1:
                    subject = user.subjects[subject_index]
                    result += '{}:   {}\n'.format(index + 1, subject.name)

                    if subject.teacher is not None:
                        result += subject.teacher + '\n'

                    if subject.room is not None:
                        result += subject.room + '\n'

                    homework = session.query(Homework).filter_by(date=ordinal_date, subject=subject_index,
                                                                 for_lesson=True, user_id=user.id).all()
                    for h in homework:
                        result += '❗<b>{}</b>\n'.format(h.description)

                    result += '\n'

    homework = session.query(Homework).filter_by(date=ordinal_date, for_lesson=False, user_id=user.id) \
        .order_by(Homework.subject).all()
    subject_index = None
    for h in homework:
        if subject_index != h.subject:
            subject_index = h.subject
            subject_name = user.subjects[subject_index].name
            result += '\n' + subject_name + '\n'
        result += '❗<b>{}</b>\n'.format(h.description)

    session.close()

    return result


def in_schedule(user_id, ordinal_date, subject_index):
    session, user = get_user(user_id)
    if not user.schedule:
        session.close()
        raise ScheduleNotFoundError

    date = datetime.date.fromordinal(ordinal_date)
    week = date.isocalendar()[1] % 2
    week_day = date.weekday()
    if week_day == 6:
        return False

    schedule = json.loads(user.schedule)
    session.close()

    return subject_index in schedule[week][week_day]


def get_subjects(user_id):
    session, user = get_user(user_id)
    if not user.subjects:
        session.close()
        raise ScheduleNotFoundError

    subjects = [s.name for s in user.subjects]
    session.close()

    return subjects


def add_homework(user_id, subject_index, date, for_lesson, description):
    session, user = get_user(user_id)
    if not user.schedule:
        session.close()
        raise ScheduleNotFoundError

    if not date:
        schedule = json.loads(user.schedule)

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
                break

            date += 1
            week_day += 1

    user.homework.append(Homework(date=date, subject=subject_index, for_lesson=for_lesson, description=description))

    session.commit()
    session.close()

    return date


def get_dates(user_id, ordinal=False):
    session, user = get_user(user_id)

    if not user.homework:
        session.close()
        return None

    today = datetime.datetime.now(TIMEZONE).date().toordinal()

    ordinal_dates = [h.date for h in user.homework if h.date >= today]
    session.close()

    if not ordinal_dates:
        return None

    ordinal_dates = sorted(list(set(ordinal_dates)))

    if ordinal:
        return ordinal_dates

    dates = []
    for ordinal_date in ordinal_dates:
        if ordinal_date == today:
            dates.append('Today')
        elif ordinal_date == today + 1:
            dates.append('Tomorrow')
        else:
            date = datetime.date.fromordinal(ordinal_date)
            dates.append('{}.{} ({})'.format(str(date.day).zfill(2), str(date.month).zfill(2),
                                             SHORT_WEEK_DAYS[date.weekday()]))

    return dates


def get_homework(user_id, ordinal_date):
    try:
        subjects = get_subjects(user_id)
    except ScheduleNotFoundError:
        return None

    session, user = get_user(user_id)
    homework = session.query(Homework).filter_by(date=ordinal_date, user_id=user.id).order_by(Homework.subject).all()
    session.close()

    if not homework:
        return None

    result = []
    for h in homework:
        result.append('{} ({}):   {}'.format(subjects[h.subject], 'lesson' if h.for_lesson else 'day', h.description))

    return result


def delete_homework(user_id, ordinal_date, homework_str):
    subjects = get_subjects(user_id)

    homework_str = homework_str.split('):   ')
    description = homework_str[-1]
    homework_type = homework_str[0].split('(')[-1]
    for_lesson = True if homework_type == 'lesson' else False
    subject = homework_str[0].split(' ({}'.format(homework_type))[0]
    subject_index = subjects.index(subject)

    session, user = get_user(user_id)
    deleted_rows = session.query(Homework).filter_by(date=ordinal_date, subject=subject_index, for_lesson=for_lesson,
                                                     description=description, user_id=user.id).delete()

    session.commit()
    session.close()

    if deleted_rows == 0:
        raise Exception('Could not find the homework to delete')


def delete_past_homework():
    today = datetime.datetime.now(TIMEZONE).date().toordinal()

    session = Session()
    session.query(Homework).filter(Homework.date < today - 1).delete()

    session.commit()
    session.close()


def get_notifications():
    today = datetime.datetime.now(TIMEZONE).date().toordinal()

    session = Session()
    users = [(user.id, user.telegram_id) for user in session.query(User).all()]

    result = []
    for user_id, telegram_id in users:
        try:
            subjects = get_subjects(telegram_id)
        except ScheduleNotFoundError:
            continue

        text = ''
        today_homework = session.query(Homework).filter_by(user_id=user_id, date=today, for_lesson=False) \
            .order_by(Homework.subject).all()
        if today_homework:
            text += '<i>Today (until the end of the day):</i>\n'
            subject = None
            for h in today_homework:
                if subject != h.subject:
                    subject = h.subject
                    text += '\n' + subjects[subject] + '\n'
                text += '❗<b>{}</b>\n'.format(h.description)
            text += '\n\n'

        tomorrow_homework = session.query(Homework).filter_by(user_id=user_id, date=today + 1, for_lesson=True) \
            .order_by(Homework.subject).all()
        if tomorrow_homework:
            text += '<i>Tomorrow (only for the lessons):</i>\n'
            subject = None
            for h in tomorrow_homework:
                if subject != h.subject:
                    subject = h.subject
                    text += '\n' + subjects[subject] + '\n'
                text += '❗<b>{}</b>\n'.format(h.description)

        if text != '':
            result.append((telegram_id, text))

    session.close()

    return result
