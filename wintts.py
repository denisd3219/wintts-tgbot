import os

#pypiwin32
from pythoncom import CoInitialize 
from win32com.client import Dispatch 

#python-telegram-bot
import logging

from telegram import ReplyKeyboardMarkup, ReplyKeyboardRemove, Update, InputMediaAudio
from telegram.ext import (
    Updater,
    CommandHandler,
    MessageHandler,
    Filters,
	MessageFilter,
    ConversationHandler,
    CallbackContext,
)

logging.basicConfig(
	format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO
)
logger = logging.getLogger(__name__)

def get_voicenames():
	CoInitialize()
	sapi = Dispatch('SAPI.SpVoice')
	return [voice.GetDescription() for voice in sapi.GetVoices()]

def text_to_file(filename, msg, rate=0, vn='Microsoft David Desktop - English (United States)', volume=100):
	CoInitialize()
	sapi = Dispatch('SAPI.SpVoice')
	fs = Dispatch('SAPI.SpFileStream')
	fname = '%s.wav'%filename
	fs.Open(fname, 3)
	sapi.AudioOutputStream = fs
	sapi.Rate = max(0, min(10, int(rate)))
	sapi.Volume = max(0, min(100, int(volume)))
	voices = sapi.GetVoices()
	if vn in get_voicenames():
		for v in voices:
			if vn in v.GetDescription():
				sapi.Voice = v
	else:
		if vn in range(len(voices)):
			sapi.Voice = voices[vn]
	sapi.Speak(msg)
	fs.Close()
	return fname

def get_voicestring():
	CoInitialize()
	sapi = Dispatch('SAPI.SpVoice')
	voices = sapi.GetVoices()
	vstr = ""
	for index in range(len(voices)):
		vstr += "%s | %s \n" % (index, voices[index].GetDescription())
	return vstr

class FilterVoice(MessageFilter):
    def filter(self, message):
        return message.text in get_voicenames()

filter_voice = FilterVoice()


VOICE, RATE, VOLUME, MSGTYPE, MSG = range(5)

def start(update: Update, context: CallbackContext) -> int:
	reply_keyboard = [get_voicenames()]
	update.message.reply_text(
		'Choose a voice',
		reply_markup=ReplyKeyboardMarkup(reply_keyboard, one_time_keyboard=True),
	)

	return VOICE

def voice(update: Update, context: CallbackContext) -> int:
	user = update.message.from_user
	context.user_data['voice'] = update.message.text

	update.message.reply_text(
		'Set speaking rate (0-10)',
		reply_markup=ReplyKeyboardRemove(),
	)

	return RATE

def rate(update: Update, context: CallbackContext) -> int:
	user = update.message.from_user
	context.user_data['rate'] = update.message.text

	update.message.reply_text(
		'Set volume (0-100)'
	)

	return VOLUME

def volume(update: Update, context: CallbackContext) -> int:
	user = update.message.from_user
	context.user_data['volume'] = update.message.text

	reply_keyboard = [['Audio', 'Voice']]
	update.message.reply_text(
		'Choose an output message type',
		reply_markup=ReplyKeyboardMarkup(reply_keyboard, one_time_keyboard=True),
	)

	return MSGTYPE

def msgtype(update: Update, context: CallbackContext) -> int:
	user = update.message.from_user
	context.user_data['msgtype'] = update.message.text

	update.message.reply_text(
		'Write a text to say',
		reply_markup=ReplyKeyboardRemove(),
	)

	return MSG

def msg(update: Update, context: CallbackContext) -> int:
	user = update.message.from_user
	fn = text_to_file(user.first_name, update.message.text, context.user_data['rate'], context.user_data['voice'], context.user_data['volume'])
	f = open(fn, 'rb')

	if context.user_data['msgtype'] == 'Audio':
		update.message.reply_audio(f)
	else:
		update.message.reply_voice(f)

	f.close()
	os.remove(fn)

	return ConversationHandler.END

def cancel(update: Update, context: CallbackContext) -> int:
    user = update.message.from_user
    update.message.reply_text(
        'Canceled', reply_markup=ReplyKeyboardRemove()
    )

    return ConversationHandler.END

def main():
	updater = Updater("TOKEN")

	dispatcher = updater.dispatcher

	conv_handler = ConversationHandler(
		entry_points=[CommandHandler('start', start)],
		states={
			VOICE: [MessageHandler(filter_voice, voice)],
			RATE: [MessageHandler(Filters.text, rate)],
			VOLUME: [MessageHandler(Filters.text, volume)],
			MSGTYPE: [MessageHandler(Filters.regex('^(Audio|Voice)$'), msgtype)],
			MSG: [MessageHandler(Filters.text, msg)],
		},
		fallbacks=[CommandHandler('cancel', cancel)],
	)

	dispatcher.add_handler(conv_handler)
	updater.start_polling()
	updater.idle()

if __name__ == '__main__':
	main()