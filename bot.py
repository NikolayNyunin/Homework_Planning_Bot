import time
import datetime
from threading import Thread

import telebot
from telebot.types import ReplyKeyboardMarkup, ReplyKeyboardRemove
import schedule

from my_token import TOKEN
from planning import TIMEZONE, set_schedule, get_schedule, in_schedule, get_subjects, add_homework, get_dates,\
    get_homework, delete_homework, delete_past_homework, get_notifications

MARKUP = ReplyKeyboardMarkup(resize_keyboard=True).add('Today', 'Tomorrow', 'Week').add('Add', 'Delete').add('Info')

bot = telebot.TeleBot(TOKEN)
data = {}


def check_cancel(message, adding=True):
    if message.text.lower() in ('cancel', '❌ cancel ❌'):
        if message.from_user.id in data:
            del data[message.from_user.id]
        bot.send_message(message.chat.id, 'Homework {} was cancelled.'.format('adding' if adding else 'deleting'),
                         reply_markup=MARKUP)
        return True

    return False


def check_data(message):
    if message.from_user.id not in data:
        bot.send_message(message.chat.id, 'Something went wrong. Please try again.')
        return True

    return False


def process_date(message, function):
    try:
        text = message.text.split('.')
        date = datetime.date(year=datetime.datetime.now(TIMEZONE).year,
                             month=int(text[1]), day=int(text[0])).toordinal()
        if date < datetime.datetime.now(TIMEZONE).date().toordinal():
            bot.send_message(message.chat.id, 'Error: Past date.\n'
                                              'Please enter a future or present date.')
            bot.register_next_step_handler(message, function)
        else:
            return date

    except Exception as e:
        bot.send_message(message.chat.id, 'Error: Incorrect date ({}).\n'
                                          'Make sure to type it in DD.MM format.'.format(str(e)))
        bot.register_next_step_handler(message, function)


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
                                      'To add new homework, press "Add" and follow instructions.\n'
                                      'To delete existing homework, press "Delete".\n'
                                      'You can cancel adding or deleting homework by typing "Cancel" or pressing '
                                      'the corresponding button.', reply_markup=MARKUP)


@bot.message_handler(content_types=['document'])
def handle_document(message):
    if not message.document.file_name.endswith('.xlsx'):
        bot.send_message(message.chat.id, 'Error: Unsupported file type.')
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
        bot.send_message(message.chat.id, 'Error: Incorrect response.\nPlease choose one of the options.')
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

        elif text == 'add':
            subjects = get_subjects(message.from_user.id)
            markup = ReplyKeyboardMarkup(resize_keyboard=True).add(*subjects, row_width=1).add('❌ Cancel ❌')
            bot.send_message(message.chat.id, 'Choose the subject.', reply_markup=markup)
            bot.register_next_step_handler(message, handle_subject)

        elif text == 'delete':
            dates = get_dates(message.from_user.id)
            if not dates:
                bot.send_message(message.chat.id, 'You have no recorded homework to delete.')
                return

            markup = ReplyKeyboardMarkup(resize_keyboard=True).add(*dates, row_width=3).add('❌ Cancel ❌')
            bot.send_message(message.chat.id, 'Choose the date of the homework you wish to delete '
                                              '(or type it as DD.MM)', reply_markup=markup)
            bot.register_next_step_handler(message, handle_existing_date)

        else:
            bot.send_message(message.chat.id, "Bot didn't understand you.", reply_markup=MARKUP)

    except FileNotFoundError:
        bot.send_message(message.chat.id, 'Error: Schedule not found.\n'
                                          'Please set your schedule before requesting it.')
    except Exception as e:
        bot.send_message(message.chat.id, 'Error: {}.'.format(str(e)))


def handle_subject(message):
    try:
        if message.text is None:
            bot.send_message(message.chat.id, 'Error: Empty response.\nPlease choose one of the options.')
            bot.register_next_step_handler(message, handle_subject)
            return

        elif check_cancel(message):
            return

        subjects = get_subjects(message.from_user.id)

        if message.text not in subjects:
            bot.send_message(message.chat.id, 'Error: Incorrect response.\nPlease choose one of the options.')
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

        markup.add(*dates, row_width=4).add('❌ Cancel ❌')
        bot.send_message(message.chat.id, 'Choose the deadline: press one of the buttons '
                                          'or type your own date in DD.MM format.', reply_markup=markup)
        bot.register_next_step_handler(message, handle_new_date)

    except FileNotFoundError:
        bot.send_message(message.chat.id, 'Error: Schedule not found.\n'
                                          'Please set your schedule before requesting it.')
    except Exception as e:
        bot.send_message(message.chat.id, 'Error: {}.'.format(str(e)))


def handle_new_date(message):
    try:
        if message.text is None:
            bot.send_message(message.chat.id, 'Error: Empty date message.')
            bot.register_next_step_handler(message, handle_new_date)
            return

        elif check_cancel(message) or check_data(message):
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
            date = process_date(message, handle_new_date)
            if not date:
                return

            data[message.from_user.id].append(date)

        subject, date = data[message.from_user.id]
        if date and in_schedule(message.from_user.id, date, subject):
            bot.send_message(message.chat.id, 'Is this deadline set for the lesson or the end of the day?',
                             reply_markup=ReplyKeyboardMarkup(resize_keyboard=True)
                             .add('Lesson', 'Day').add('❌ Cancel ❌'))
            bot.register_next_step_handler(message, handle_type)
            return
        elif date:
            data[message.from_user.id].append(False)
        else:
            data[message.from_user.id].append(True)

        bot.send_message(message.chat.id, 'Write homework description. Type "Cancel" to cancel adding the homework.',
                         reply_markup=ReplyKeyboardRemove())
        bot.register_next_step_handler(message, handle_description)

    except FileNotFoundError:
        bot.send_message(message.chat.id, 'Error: Schedule not found.\n'
                                          'Please set your schedule before requesting it.')
    except Exception as e:
        bot.send_message(message.chat.id, 'Error: {}.'.format(str(e)))


def handle_type(message):
    try:
        if message.text is None or message.text.lower() not in ('lesson', 'day', 'cancel', '❌ cancel ❌'):
            bot.send_message(message.chat.id, 'Error: Incorrect response.\nPlease choose one of the options.')
            bot.register_next_step_handler(message, handle_type)
            return

        elif check_cancel(message) or check_data(message):
            return

        text = message.text.lower()
        if text == 'lesson':
            data[message.from_user.id].append(True)
        elif text == 'day':
            data[message.from_user.id].append(False)

        bot.send_message(message.chat.id, 'Write homework description. Type "Cancel" to cancel adding the homework.',
                         reply_markup=ReplyKeyboardRemove())
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

        elif check_cancel(message) or check_data(message):
            return

        subject, date, for_lesson = data.pop(message.from_user.id)
        add_homework(message.from_user.id, subject, date, for_lesson, message.text)
        bot.send_message(message.chat.id, 'Homework successfully added.', reply_markup=MARKUP)

    except FileNotFoundError:
        bot.send_message(message.chat.id, 'Error: Schedule not found.\n'
                                          'Please set your schedule before requesting it.')
    except Exception as e:
        bot.send_message(message.chat.id, 'Error: {}.'.format(str(e)))


def handle_existing_date(message):
    if message.text is None:
        bot.send_message(message.chat.id, 'Error: Empty date message.')
        bot.register_next_step_handler(message, handle_existing_date)
        return

    elif check_cancel(message, adding=False):
        return

    text = message.text.lower()
    if text == 'today':
        date = datetime.datetime.now(TIMEZONE).date().toordinal()
    elif text == 'tomorrow':
        date = datetime.datetime.now(TIMEZONE).date().toordinal() + 1
    else:
        date = process_date(message, handle_existing_date)
        if not date:
            return

    homework = get_homework(message.from_user.id, date)
    if not homework:
        bot.send_message(message.chat.id, 'You have no recorded homework for that date.\nChoose the date again.')
        bot.register_next_step_handler(message, handle_existing_date)
        return

    data[message.from_user.id] = date
    markup = ReplyKeyboardMarkup(resize_keyboard=True).add(*homework, row_width=1).add('❌ Cancel ❌')
    bot.send_message(message.chat.id, 'Choose the homework to delete.', reply_markup=markup)
    bot.register_next_step_handler(message, handle_homework)


def handle_homework(message):
    if message.text is None:
        bot.send_message(message.chat.id, 'Error: Empty homework message.')
        bot.register_next_step_handler(message, handle_homework)
        return

    elif check_cancel(message, adding=False) or check_data(message):
        return

    try:
        delete_homework(message.from_user.id, data[message.from_user.id], message.text)
        bot.send_message(message.chat.id, 'Homework deleted successfully.', reply_markup=MARKUP)
        del data[message.from_user.id]

    except Exception as e:
        bot.send_message(message.chat.id, 'Error: {}.\n'
                                          'Choose the homework you want to delete again.'.format(e))
        bot.register_next_step_handler(message, handle_homework)


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
