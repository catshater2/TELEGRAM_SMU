import logging
import pandas as pd
import re
import requests
from datetime import datetime
from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import StatesGroup, State

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

EXCEL_FILE = 'user_data.xlsx'

class Form(StatesGroup):
    location = State()
    fio = State()
    bs_number = State()
    stages = State()

def get_city_by_coordinates(shirota, dolgota):
    url = f"https://nominatim.openstreetmap.org/reverse?lat={shirota}&lon={dolgota}&format=json&accept-language=ru"
    headers = {'User-Agent': 'MyBot/1.0 (contact@example.com)'}
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        data = response.json()
        address = data.get("address", {})
        return address.get('city', address.get('town', address.get('village', 'Локация не определена')))
    except Exception as e:
        logger.error(f"Ошибка: {e}")
        return "Ошибка геолокации"

def save_to_excel(user_data: dict, user_id: int):
    current_date = datetime.now().strftime("%d.%m.%Y %H:%M")
    new_data = pd.DataFrame([{
        'Дата': current_date,
        'user_id': user_id,
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
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    keyboard.add(types.KeyboardButton(text="Отправить геолокацию", request_location=True))
    await message.answer("Отправьте вашу геолокацию:", reply_markup=keyboard)

async def process_location(message: types.Message, state: FSMContext):
    if not message.location:
        await message.answer("Используйте кнопку для отправки геолокации.")
        return
    
    shirota = message.location.latitude
    dolgota = message.location.longitude
    city = get_city_by_coordinates(shirota, dolgota)
    
    await state.update_data(region=city, city=city, location=(shirota, dolgota))
    await state.set_state(Form.fio)
    await message.answer(f"Широта: {shirota}, Долгота: {dolgota}\nГород: {city}\n\nТеперь введите ваше ФИО:")

async def process_fio(message: types.Message, state: FSMContext):
    await state.update_data(fio=message.text)
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    keyboard.add("Подтвердить", "Изменить")
    await message.answer(f"Проверьте ФИО: {message.text}", reply_markup=keyboard)
    await state.set_state(Form.bs_number)

async def station_number(message: types.Message, state: FSMContext):
    if message.text == 'Изменить':
        await state.set_state(Form.fio)
        await message.answer("Введите ФИО:")
        return
    
    text = message.text.upper().strip()
    if not re.match(r'^[A-Z]{2}-\d{6}$', text):
        await message.answer("Неверный формат")
        return
    
    await state.update_data(bs_number=text)
    await state.set_state(Form.stages)
    
    keyboard = types.InlineKeyboardMarkup(row_width=1)
    stages = ['Этап 1', 'Этап 2', 'Этап 3', 'Этап 4', 'Этап 5', 'Этап 6', 'Этап 7', 'Этап 8', 'Этап 9']
    for stage in stages:
        keyboard.add(types.InlineKeyboardButton(text=f"{stage} ❌", callback_data=f"stage_{stage}"))
    keyboard.add(types.InlineKeyboardButton(text="✅ Сохранить", callback_data="done"))
    
    await message.answer("Выберите этапы:", reply_markup=keyboard)

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
        
        keyboard = types.InlineKeyboardMarkup(row_width=1)
        stages = ['Этап 1', 'Этап 2', 'Этап 3', 'Этап 4', 'Этап 5', 'Этап 6', 'Этап 7', 'Этап 8', 'Этап 9']
        for s in stages:
            status = '✅' if s in selected else '❌'
            keyboard.add(types.InlineKeyboardButton(text=f"{s} {status}", callback_data=f"stage_{s}"))
        keyboard.add(types.InlineKeyboardButton(text="✅ Сохранить", callback_data="done"))
        
        await callback.message.edit_reply_markup(reply_markup=keyboard)
        await callback.answer()
    
    elif callback.data == 'done':
        if not selected:
            await callback.answer("Выберите хотя бы один этап", show_alert=True)
            return
        
        user_data = await state.get_data()
        save_to_excel(user_data, callback.from_user.id)
        await callback.message.edit_text("Данные сохранены")
        await state.clear()
        await callback.answer()

async def send_excel_handler(message: types.Message):
    try:  
        with open(EXCEL_FILE, 'rb') as file: 
            await message.answer_document(file, caption='Актуальные данные')
    except FileNotFoundError:
        await message.answer("Не создан файл")

async def cancel_handler(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer('Сессия завершена')

async def main():
    bot = Bot(token="7856193785:AAFhbj0B8TI33LuXJnwA8PimgwM07PauAg8")
    dp = Dispatcher()
    
    dp.message.register(start_handler, Command('start'))
    dp.message.register(send_excel_handler, Command('get_excel'))
    dp.message.register(cancel_handler, Command('cancel'))
    
    dp.message.register(process_location, Form.location)
    dp.message.register(process_fio, Form.fio)
    dp.message.register(station_number, Form.bs_number)
    
    dp.callback_query.register(process_stages, Form.stages)
    
    await dp.start_polling(bot)

if __name__ == '__main__':
    import asyncio
    asyncio.run(main())