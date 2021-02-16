import telebot
from telebot import types

from my_token import TOKEN
from schedule import set_schedule

bot = telebot.TeleBot(TOKEN)


@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, 'Welcome to Homework Planning Bot!\n'
                                      'It can monitor all your homework and notify you about your deadlines.\n'
                                      'For more information type /info.\n')


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
    bot.send_message(message.chat.id, 'Are you sure you want to change your schedule?\n'
                                      'This will delete all your recorded homework.\n'
                                      'Answer "Yes" or "No".',
                     reply_markup=types.ForceReply())
    bot.register_next_step_handler(message, handle_answer)


def handle_answer(message):
    if message.text is None:
        bot.send_message(message.chat.id, "Incorrect response. Schedule wasn't changed.")
    elif message.text.lower() == 'yes':
        path = 'files/' + str(message.from_user.id) + '.xlsx'
        try:
            with open(path, 'rb') as file:
                set_schedule(message.from_user.id, file)
        except Exception as e:
            bot.send_message(message.chat.id, 'Error: ' + str(e))
        else:
            bot.send_message(message.chat.id, 'New schedule successfully set.')
    elif message.text.lower() == 'no':
        bot.send_message(message.chat.id, "Schedule wasn't changed.")
    else:
        bot.send_message(message.chat.id, "Bot couldn't understand you. Please type 'Yes' or 'No'.",
                         reply_markup=types.ForceReply())
        bot.register_next_step_handler(message, handle_answer)


bot.polling()
