import time
from threading import Thread

import telebot
from telebot.types import ReplyKeyboardMarkup, ReplyKeyboardRemove
import schedule

from my_token import TOKEN
from planning import *

MARKUP = ReplyKeyboardMarkup(resize_keyboard=True).add('Today', 'Tomorrow', 'Week').add('Homework').add('Info')

bot = telebot.TeleBot(TOKEN)
data = {}


@bot.message_handler(commands=['start'])
@bot.message_handler(func=lambda message: message.text is not None and message.text.lower() == 'start')
def start(message):
    bot.send_message(message.chat.id, 'Welcome to Homework Planning Bot!\n'
                                      'Type /info or /help to learn what it can do.', reply_markup=MARKUP)


@bot.message_handler(commands=['info', 'help'])
@bot.message_handler(func=lambda message: message.text is not None and message.text.lower() in ('info', 'help'))
def info(message):
    bot.send_message(message.chat.id, 'This is Homework Planning Bot.\n'
                                      'It can monitor your homework.\n'
                                      'To set your schedule, attach .xlsx file with the specific structure.\n'
                                      'To view your schedule and homework, press "Today", "Tomorrow" or "Week".\n'
                                      'To add new homework, press "Homework" and follow instructions.',
                     reply_markup=MARKUP)


@bot.message_handler(content_types=['document'])
def handle_document(message):
    if not message.document.file_name.endswith('.xlsx'):
        bot.send_message(message.chat.id, 'Unsupported file type.')
        return

    file_info = bot.get_file(message.document.file_id)
    file = bot.download_file(file_info.file_path)
    path = 'files/' + str(message.from_user.id) + '.xlsx'
    with open(path, 'wb') as new_file:
        new_file.write(file)

    markup = ReplyKeyboardMarkup(resize_keyboard=True).add('Yes', 'No')
    bot.send_message(message.chat.id, 'Are you sure you want to change your schedule?\n'
                                      'This will delete all your recorded homework.', reply_markup=markup)
    bot.register_next_step_handler(message, handle_change_schedule_answer)


def handle_change_schedule_answer(message):
    if message.text is None or message.text.lower() not in ('yes', 'no'):
        bot.send_message(message.chat.id, 'Incorrect response. Please choose one of the options.')
        bot.register_next_step_handler(message, handle_change_schedule_answer)

    elif message.text.lower() == 'yes':
        path = 'files/' + str(message.from_user.id) + '.xlsx'
        try:
            with open(path, 'rb') as file:
                set_schedule(message.from_user.id, file)
        except Exception as e:
            bot.send_message(message.chat.id, 'Error: {}.'.format(str(e)))
        else:
            bot.send_message(message.chat.id, 'New schedule successfully set.', reply_markup=MARKUP)

    elif message.text.lower() == 'no':
        bot.send_message(message.chat.id, "Schedule wasn't changed.", reply_markup=MARKUP)


@bot.message_handler(content_types=['text'])
def handle_text(message):
    try:
        text = message.text.lower()
        if text == 'today':
            date = datetime.datetime.now(TIMEZONE).date().toordinal()
            bot.send_message(message.chat.id, get_schedule(message.from_user.id, date), parse_mode='HTML')

        elif text == 'tomorrow':
            date = datetime.datetime.now(TIMEZONE).date().toordinal() + 1
            bot.send_message(message.chat.id, get_schedule(message.from_user.id, date), parse_mode='HTML')

        elif text == 'week':
            date = datetime.datetime.now(TIMEZONE).date().toordinal()
            for i in range(7):
                bot.send_message(message.chat.id, get_schedule(message.from_user.id, date), parse_mode='HTML')
                date += 1

        elif text == 'homework':
            subjects = get_subjects(message.from_user.id)
            markup = ReplyKeyboardMarkup(resize_keyboard=True).add(*subjects, row_width=1)
            bot.send_message(message.chat.id, 'Choose the subject.', reply_markup=markup)
            bot.register_next_step_handler(message, handle_subject)

        else:
            bot.send_message(message.chat.id, "Bot couldn't understand you.", reply_markup=MARKUP)

    except FileNotFoundError:
        bot.send_message(message.chat.id, 'Error: Schedule not found.\n'
                                          'Please set your schedule before requesting it.')
    except Exception as e:
        bot.send_message(message.chat.id, 'Error: {}.'.format(str(e)))


def handle_subject(message):
    try:
        subjects = get_subjects(message.from_user.id)

        if message.text is None or message.text not in subjects:
            bot.send_message(message.chat.id, 'Incorrect response. Please choose one of the options.')
            bot.register_next_step_handler(message, handle_subject)
            return

        index = subjects.index(message.text)
        data[message.from_user.id] = [index]

        ordinal_date = datetime.datetime.now(TIMEZONE).date().toordinal() + 1
        markup = ReplyKeyboardMarkup(resize_keyboard=True).add('Next lesson').add('Today', 'Tomorrow')
        dates = []
        for i in range(12):
            ordinal_date += 1
            date = datetime.date.fromordinal(ordinal_date)
            dates.append('{}.{}'.format(str(date.day).zfill(2), str(date.month).zfill(2)))

        markup.add(*dates, row_width=4)
        bot.send_message(message.chat.id, 'Choose the deadline: press one of the buttons '
                                          'or type your own date in DD.MM format.', reply_markup=markup)
        bot.register_next_step_handler(message, handle_date)

    except FileNotFoundError:
        bot.send_message(message.chat.id, 'Error: Schedule not found.\n'
                                          'Please set your schedule before requesting it.')
    except Exception as e:
        bot.send_message(message.chat.id, 'Error: {}.'.format(str(e)))


def handle_date(message):
    try:
        if message.text is None:
            bot.send_message(message.chat.id, 'Error: Empty date message.')
            bot.register_next_step_handler(message, handle_date)
            return

        elif message.from_user.id not in data:
            subjects = get_subjects(message.from_user.id)
            markup = ReplyKeyboardMarkup(resize_keyboard=True).add(*subjects, row_width=1)
            bot.send_message(message.chat.id, 'Something went wrong. Please choose the subject again.',
                             reply_markup=markup)
            bot.register_next_step_handler(message, handle_subject)
            return

        text = message.text.lower()
        if text == 'next lesson':
            data[message.from_user.id].append(None)
        elif text == 'today':
            date = datetime.datetime.now(TIMEZONE).date().toordinal()
            data[message.from_user.id].append(date)
        elif text == 'tomorrow':
            date = datetime.datetime.now(TIMEZONE).date().toordinal() + 1
            data[message.from_user.id].append(date)
        else:
            try:
                text = text.split('.')
                ordinal_date = datetime.date(year=datetime.datetime.now(TIMEZONE).year,
                                             month=int(text[1]), day=int(text[0])).toordinal()
            except Exception as e:
                bot.send_message(message.chat.id, 'Error: Incorrect date ({}).\n'
                                                  'Make sure to type it in DD.MM format.'.format(str(e)))
                bot.register_next_step_handler(message, handle_date)
                return

            data[message.from_user.id].append(ordinal_date)

        subject, date = data[message.from_user.id]
        if date and in_schedule(message.from_user.id, date, subject):
            bot.send_message(message.chat.id, 'Is this deadline set for the lesson or the end of the day?',
                             reply_markup=ReplyKeyboardMarkup(resize_keyboard=True).add('Lesson', 'Day'))
            bot.register_next_step_handler(message, handle_type)
            return
        elif date:
            data[message.from_user.id].append(False)
        else:
            data[message.from_user.id].append(True)

        bot.send_message(message.chat.id, 'Write homework description.', reply_markup=ReplyKeyboardRemove())
        bot.register_next_step_handler(message, handle_description)

    except FileNotFoundError:
        bot.send_message(message.chat.id, 'Error: Schedule not found.\n'
                                          'Please set your schedule before requesting it.')
    except Exception as e:
        bot.send_message(message.chat.id, 'Error: {}.'.format(str(e)))


def handle_type(message):
    try:
        if message.text is None or message.text.lower() not in ('lesson', 'day'):
            bot.send_message(message.chat.id, 'Incorrect response. Please choose one of the options.')
            bot.register_next_step_handler(message, handle_type)
            return

        elif message.from_user.id not in data:
            subjects = get_subjects(message.from_user.id)
            markup = ReplyKeyboardMarkup(resize_keyboard=True).add(*subjects, row_width=1)
            bot.send_message(message.chat.id, 'Something went wrong. Please choose the subject again.',
                             reply_markup=markup)
            bot.register_next_step_handler(message, handle_subject)
            return

        text = message.text.lower()
        if text == 'lesson':
            data[message.from_user.id].append(True)
        elif text == 'day':
            data[message.from_user.id].append(False)

        bot.send_message(message.chat.id, 'Write homework description.', reply_markup=ReplyKeyboardRemove())
        bot.register_next_step_handler(message, handle_description)

    except FileNotFoundError:
        bot.send_message(message.chat.id, 'Error: Schedule not found.\n'
                                          'Please set your schedule before requesting it.')
    except Exception as e:
        bot.send_message(message.chat.id, 'Error: {}.'.format(str(e)))


def handle_description(message):
    try:
        if message.text is None:
            bot.send_message(message.chat.id, 'Error: Empty description.')
            bot.register_next_step_handler(message, handle_description)
            return

        elif message.from_user.id not in data:
            subjects = get_subjects(message.from_user.id)
            markup = ReplyKeyboardMarkup(resize_keyboard=True).add(*subjects, row_width=1)
            bot.send_message(message.chat.id, 'Something went wrong. Please choose the subject again.',
                             reply_markup=markup)
            bot.register_next_step_handler(message, handle_subject)
            return

        subject, date, for_lesson = data.pop(message.from_user.id)
        add_homework(message.from_user.id, subject, date, for_lesson, message.text)
        bot.send_message(message.chat.id, 'Homework successfully added.', reply_markup=MARKUP)

    except FileNotFoundError:
        bot.send_message(message.chat.id, 'Error: Schedule not found.\n'
                                          'Please set your schedule before requesting it.')
    except Exception as e:
        bot.send_message(message.chat.id, 'Error: {}.'.format(str(e)))


def send_notifications():
    notifications = get_notifications()
    if not notifications:
        return

    for n in notifications:
        bot.send_message(n[0], n[1], parse_mode='HTML')


def main():
    bot.polling(none_stop=True)


def timer():
    schedule.every().day.do(delete_past_homework)
    schedule.every().day.at('18:00').do(send_notifications)
    while True:
        schedule.run_pending()
        time.sleep(1)


if __name__ == '__main__':
    bot_thread = Thread(target=main, name='BotThread')
    timer_thread = Thread(target=timer, name='TimerThread')

    bot_thread.start()
    timer_thread.start()
