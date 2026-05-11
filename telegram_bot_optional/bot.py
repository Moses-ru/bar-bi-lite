import asyncio
import os
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, KeyboardButton, ReplyKeyboardMarkup, WebAppInfo
from aiogram.filters import CommandStart
from dotenv import load_dotenv

load_dotenv()
TOKEN = os.getenv('TELEGRAM_BOT_TOKEN')
WEBAPP_URL = os.getenv('WEBAPP_URL')

if not TOKEN:
    raise RuntimeError('Добавь TELEGRAM_BOT_TOKEN в .env')
if not WEBAPP_URL:
    raise RuntimeError('Добавь WEBAPP_URL в .env')

bot = Bot(TOKEN)
dp = Dispatcher()

@dp.message(CommandStart())
async def start(message: Message):
    kb = ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text='Открыть Бар BI', web_app=WebAppInfo(url=WEBAPP_URL))]],
        resize_keyboard=True
    )
    await message.answer('Открой мобильный Бар BI WebApp:', reply_markup=kb)

async def main():
    print('Бот запущен')
    await dp.start_polling(bot)

if __name__ == '__main__':
    asyncio.run(main())
