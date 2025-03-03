import logging
import pandas as pd
import re
import requests
from datetime import datetime
from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import StatesGroup, State
from aiogram.utils.keyboard import InlineKeyboardBuilder
from dotenv import load_dotenv
import os

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

EXCEL_FILE = 'user_data.xlsx'

class Form(StatesGroup):
    location = State()
    fio = State()
    confirm_fio = State()
    bs_number = State()
    stages = State()
def get_city_by_coordinates(latitude, longitude):
    url = f"https://nominatim.openstreetmap.org/reverse?lat={latitude}&lon={longitude}&format=json&accept-language=ru"
    headers = {'User-Agent': 'MyBot/1.0 (contact@example.com)'}
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        data = response.json()
        address = data.get("address", {})
        
        # Приоритетно ищем город
        city = (address.get('city') or 
                address.get('town') or 
                address.get('village') or 
                address.get('municipality') or 
                address.get('hamlet') or 
                address.get('locality'))
        
        # Если город не найден, возвращаем более общее значение (например, район или регион)
        if not city:
            city = (address.get('county') or 
                    address.get('state_district') or 
                    address.get('region') or 
                    address.get('state'))
        
        # Если всё ещё не найдено, возвращаем "Локация не определена"
        if not city:
            logger.warning(f"Локация не найдена, полный адрес: {address}")
            return "Локация не определена"
        
        return city
    except Exception as e:
        logger.error(f"Ошибка: {e}")
        return "Ошибка геолокации"

def save_to_excel(user_data: dict, user_id: int):
    current_date = datetime.now().strftime("%d.%m.%Y %H:%M")
    new_data = pd.DataFrame([{
        'Дата': current_date,
        'Регион': user_data.get('region'),
        'Город': user_data.get('city', 'Не определен'),
        'ФИО': user_data.get('fio'),
        'Номер сайта': user_data.get('bs_number'),
        'Этап': ', '.join(user_data.get('selected_stages', []))
    }])
    
    try:
        df = pd.read_excel(EXCEL_FILE)
        df = pd.concat([df, new_data], ignore_index=True)
    except FileNotFoundError:
        df = new_data
    
    try:
        df.to_excel(EXCEL_FILE, index=False)
    except Exception as e:
        logger.error(f"Excel: {e}")

async def start_handler(message: types.Message, state: FSMContext):
    await state.set_state(Form.location)
    
    keyboard = types.ReplyKeyboardMarkup(
        keyboard=[
            [types.KeyboardButton(text="Отправить геолокацию", request_location=True)]
        ],
        resize_keyboard=True,
        one_time_keyboard=True
    )
    
    await message.answer("Отправьте вашу геолокацию:", reply_markup=keyboard)

async def process_location(message: types.Message, state: FSMContext):
    if not message.location:
        await message.answer("Используйте кнопку геолокации.")
        return
    
    latitude = message.location.latitude
    longitude = message.location.longitude
    city = get_city_by_coordinates(latitude, longitude)
    
    await state.update_data(region=city, city=city, location=(latitude, longitude))
    await state.set_state(Form.fio)
    
    await message.answer(f" Координаты получены!\n Широта: {latitude}, Долгота: {longitude}\n Город: {city}\n\nТеперь введите ваше ФИО:")

async def process_fio(message: types.Message, state: FSMContext):
    fio = message.text.strip()
    
    if not re.match(r"^[А-Яа-яЁё\s-]+$", fio):
        await message.answer("⚠ Неверный формат ФИО. Введите ФИО снова (только буквы и пробелы).")
        return
    
    await state.update_data(fio=fio)
    
    keyboard = types.ReplyKeyboardMarkup(
        keyboard=[
            [types.KeyboardButton(text="Подтвердить")],
            [types.KeyboardButton(text="Изменить")]
        ],
        resize_keyboard=True,
        one_time_keyboard=True
    )
    
    await state.set_state(Form.confirm_fio)
    await message.answer(f"🔍 Проверьте ФИО: {fio}\nЕсли всё верно, нажмите 'Подтвердить'.", reply_markup=keyboard)

async def confirm_fio(message: types.Message, state: FSMContext):
    if message.text == "Изменить":
        await state.set_state(Form.fio)
        await message.answer("Введите ваше ФИО заново:")
    elif message.text == "Подтвердить":
        await state.set_state(Form.bs_number)
        await message.answer("Теперь введите номер базовой станции (например, XX123456):")
    else:
        await message.answer("Пожалуйста, выберите 'Подтвердить' или 'Изменить'.")

async def station_number(message: types.Message, state: FSMContext):
    text = message.text.upper().strip()
    if not re.match(r'^[A-Z]{2}\d{6}$', text):
        await message.answer("⚠ Неверный формат номера базовой станции. Введите заново (например, XX123456).")
        return
    
    await state.update_data(bs_number=text)
    await state.set_state(Form.stages)
    
    builder = InlineKeyboardBuilder()
    stages = [f'Этап {i}' for i in range(1, 10)]
    for stage in stages:
        builder.button(text=f"{stage} ❌", callback_data=f"stage_{stage}")
    builder.button(text="✅ Сохранить", callback_data="done")
    builder.adjust(1)
    
    await message.answer("🔘 Выберите этапы:", reply_markup=builder.as_markup())

async def process_stages(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    selected = data.get('selected_stages', [])
    
    if callback.data.startswith('stage_'):
        stage = callback.data.split('_', 1)[1]
        if stage in selected:
            selected.remove(stage)
        else:
            selected.append(stage)
        
        await state.update_data(selected_stages=selected)
        
        builder = InlineKeyboardBuilder()
        stages = [f'Этап {i}' for i in range(1, 10)]
        for s in stages:
            status = '✅' if s in selected else '❌'
            builder.button(text=f"{s} {status}", callback_data=f"stage_{s}")
        builder.button(text="✅ Сохранить", callback_data="done")
        builder.adjust(1)
        
        await callback.message.edit_reply_markup(reply_markup=builder.as_markup())
        await callback.answer()
    
    elif callback.data == 'done':
        if not selected:
            await callback.answer("Выберите хотя бы один этап", show_alert=True)
            return
        
        user_data = await state.get_data()
        save_to_excel(user_data, callback.from_user.id)
        await callback.message.edit_text("✅ Данные успешно сохранены!")
        await state.clear()
        await callback.answer()

async def send_excel_handler(message: types.Message):
    try:  
        with open(EXCEL_FILE, 'rb') as file: 
            await message.answer_document(
                types.BufferedInputFile(file.read(), filename='Данные.xlsx'),
                caption='📄 Актуальные данные'
            )
    except FileNotFoundError:
        await message.answer("Файл с данными не найден.")

async def cancel_handler(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer('❌ Сессия завершена.')

def setup_handlers(dp: Dispatcher):
    dp.message.register(start_handler, Command('start'))
    dp.message.register(send_excel_handler, Command('get_excel'))
    dp.message.register(cancel_handler, Command('cancel'))
    
    dp.message.register(process_location, Form.location)
    dp.message.register(process_fio, Form.fio)
    dp.message.register(confirm_fio, Form.confirm_fio)
    dp.message.register(station_number, Form.bs_number)
    
    dp.callback_query.register(process_stages, Form.stages)

async def main():
    load_dotenv()
    bot_token = os.getenv("BOT_TOKEN")
    
    if not bot_token:
        raise ValueError("BOT_TOKEN неверный")
    
    bot = Bot(token=bot_token)
    dp = Dispatcher()
    setup_handlers(dp)
    await dp.start_polling(bot)

if __name__ == '__main__':
    import asyncio
    asyncio.run(main())