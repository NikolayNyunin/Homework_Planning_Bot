import time
import datetime
from threading import Thread

import telebot
from telebot.types import ReplyKeyboardMarkup, ReplyKeyboardRemove
import schedule

from my_token import TOKEN
from planning import TIMEZONE, set_schedule, get_schedule, get_subjects, add_homework, delete_past_homework

MARKUP = ReplyKeyboardMarkup(resize_keyboard=True).add('Today', 'Tomorrow', 'Week').add('Homework').add('Info')

bot = telebot.TeleBot(TOKEN)
data = {}


@bot.message_handler(commands=['start'])
@bot.message_handler(func=lambda message: message.text.lower() == 'start')
def start(message):
    bot.send_message(message.chat.id, 'Welcome to Homework Planning Bot!\n'
                                      'Type /info or /help to learn what it can do.')


@bot.message_handler(commands=['info', 'help'])
@bot.message_handler(func=lambda message: message.text.lower() in ('info', 'help'))
def info(message):
    bot.send_message(message.chat.id, 'This is Homework Planning Bot.\n'
                                      'It can monitor your homework.\n'
                                      'To set your schedule, attach .xlsx file with the specific structure.\n'
                                      'To view your schedule and homework, press "Today", "Tomorrow" or "Week".\n'
                                      'To add new homework, press "Homework" and follow instructions.')


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

    markup = ReplyKeyboardMarkup(one_time_keyboard=True, resize_keyboard=True).add('Yes', 'No')
    bot.send_message(message.chat.id, 'Are you sure you want to change your schedule?\n'
                                      'This will delete all your recorded homework.', reply_markup=markup)
    bot.register_next_step_handler(message, handle_change_schedule_answer)


def handle_change_schedule_answer(message):
    if message.text is None or message.text.lower() not in ('yes', 'no'):
        markup = ReplyKeyboardMarkup(one_time_keyboard=True, resize_keyboard=True).add('Yes', 'No')
        bot.send_message(message.chat.id, 'Incorrect response. Please choose one of the options.', reply_markup=markup)
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
        if message.text.lower() == 'today':
            date = datetime.datetime.now(TIMEZONE).date().toordinal()
            bot.send_message(message.chat.id, get_schedule(message.from_user.id, date), parse_mode='HTML')

        elif message.text.lower() == 'tomorrow':
            date = datetime.datetime.now(TIMEZONE).date().toordinal() + 1
            bot.send_message(message.chat.id, get_schedule(message.from_user.id, date), parse_mode='HTML')

        elif message.text.lower() == 'week':
            date = datetime.datetime.now(TIMEZONE).date().toordinal()
            for i in range(7):
                bot.send_message(message.chat.id, get_schedule(message.from_user.id, date), parse_mode='HTML')
                date += 1

        elif message.text.lower() == 'homework':
            subjects = get_subjects(message.from_user.id)
            markup = ReplyKeyboardMarkup(one_time_keyboard=True, resize_keyboard=True).add(*subjects, row_width=1)
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
            markup = ReplyKeyboardMarkup(one_time_keyboard=True, resize_keyboard=True).add(*subjects, row_width=1)
            bot.send_message(message.chat.id, 'Incorrect response. Please choose one of the options.',
                             reply_markup=markup)
            bot.register_next_step_handler(message, handle_subject)
            return

        index = subjects.index(message.text)
        data[message.from_user.id] = index
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
            markup = ReplyKeyboardMarkup(one_time_keyboard=True, resize_keyboard=True).add(*subjects, row_width=1)
            bot.send_message(message.chat.id, 'Something went wrong. Please choose the subject again.',
                             reply_markup=markup)
            bot.register_next_step_handler(message, handle_subject)
            return

        subject = data.pop(message.from_user.id)
        add_homework(message.from_user.id, subject, message.text)
        bot.send_message(message.chat.id, 'Homework successfully added.', reply_markup=MARKUP)

    except FileNotFoundError:
        bot.send_message(message.chat.id, 'Error: Schedule not found.\n'
                                          'Please set your schedule before requesting it.')
    except Exception as e:
        bot.send_message(message.chat.id, 'Error: {}.'.format(str(e)))


def main():
    bot.polling(none_stop=True)


def timer():
    schedule.every().day.do(delete_past_homework)
    while True:
        schedule.run_pending()
        time.sleep(1)


if __name__ == '__main__':
    bot_thread = Thread(target=main, name='BotThread')
    timer_thread = Thread(target=timer, name='TimerThread')

    bot_thread.start()
    timer_thread.start()
