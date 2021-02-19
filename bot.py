import telebot
from telebot.types import ReplyKeyboardMarkup
import datetime

from my_token import TOKEN
from schedule import set_schedule, get_schedule

bot = telebot.TeleBot(TOKEN)


@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, 'Welcome to Homework Planning Bot!\n'
                                      'It can monitor all your homework and notify you about your deadlines.\n'
                                      'For more information type /info.')


@bot.message_handler(commands=['info'])
def info(message):
    bot.send_message(message.chat.id, 'To give bot your schedule send it Excel file.')


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

    markup = ReplyKeyboardMarkup(one_time_keyboard=True).add('Yes', 'No')
    bot.send_message(message.chat.id, 'Are you sure you want to change your schedule?\n'
                                      'This will delete all your recorded homework.', reply_markup=markup)
    bot.register_next_step_handler(message, change_schedule_answer)


def change_schedule_answer(message):
    if message.text is None or message.text.lower() not in ('yes', 'no'):
        markup = ReplyKeyboardMarkup(one_time_keyboard=True).add('Yes', 'No')
        bot.send_message(message.chat.id, "Incorrect response. Please choose one of the options.", reply_markup=markup)
        bot.register_next_step_handler(message, change_schedule_answer)

    elif message.text.lower() == 'yes':
        path = 'files/' + str(message.from_user.id) + '.xlsx'
        try:
            with open(path, 'rb') as file:
                set_schedule(message.from_user.id, file)
        except Exception as e:
            bot.send_message(message.chat.id, 'Error: ' + str(e))
        else:
            markup = ReplyKeyboardMarkup().add('Today', 'Tomorrow')
            bot.send_message(message.chat.id, 'New schedule successfully set.', reply_markup=markup)

    elif message.text.lower() == 'no':
        markup = ReplyKeyboardMarkup().add('Today', 'Tomorrow')
        bot.send_message(message.chat.id, "Schedule wasn't changed.", reply_markup=markup)


@bot.message_handler(content_types=['text'])
def handle_text(message):
    if message.text.lower() == 'today':
        week_index = datetime.date.isocalendar(datetime.date.today())[1] % 2
        day_of_the_week = datetime.date.weekday(datetime.date.today())
        try:
            bot.send_message(message.chat.id, get_schedule(message.from_user.id, week_index, day_of_the_week))
        except FileNotFoundError:
            bot.send_message(message.chat.id, 'Error: Schedule not found.\n'
                                              'Please set your schedule before requesting it.')
        except Exception as e:
            bot.send_message(message.chat.id, 'Error: ' + str(e))

    elif message.text.lower() == 'tomorrow':
        week_index = datetime.date.isocalendar(datetime.date.today())[1] % 2
        day_of_the_week = datetime.date.weekday(datetime.date.today()) + 1
        try:
            bot.send_message(message.chat.id, get_schedule(message.from_user.id, week_index, day_of_the_week))
        except FileNotFoundError:
            bot.send_message(message.chat.id, 'Error: Schedule not found.\n'
                                              'Please set your schedule before requesting it.')
        except Exception as e:
            bot.send_message(message.chat.id, 'Error: ' + str(e))

    else:
        bot.send_message(message.chat.id, "Bot couldn't understand you.")


def main():
    bot.polling()


if __name__ == '__main__':
    main()
