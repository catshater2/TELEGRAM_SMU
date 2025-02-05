import logging
import pandas as pd
import re
import requests
from datetime import datetime
from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import StatesGroup, State
from aiogram.utils.keyboard import ReplyKeyboardBuilder, InlineKeyboardBuilder

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

EXCEL_FILE = 'user_data.xlsx'

class Form(StatesGroup):
    location = State()
    fio = State()
    bs_number = State()
    stages = State()

def get_city_by_coordinates(latitude, longitude):
    """Определение города по координатам через Nominatim с расширенной логикой"""
    url = f"https://nominatim.openstreetmap.org/reverse?lat={latitude}&lon={longitude}&format=json&accept-language=ru"
    headers = {'User-Agent': 'MyBot/1.0 (contact@example.com)'}
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        data = response.json()
        
        address = data.get("address", {})
        
        location_keys = [
            'city', 'town', 'village',        # Города и поселки
            'municipality', 'hamlet',         # Муниципалитеты
            'county', 'state', 'region',      # Регионы
            'neighbourhood', 'suburb',        # Районы
            'city_district', 'road'           # Улицы
        ]
        
        for key in location_keys:
            if value := address.get(key):
                return value
        
        return "Локация не определена"
    
    except requests.exceptions.RequestException as e:
        logger.error(f"Ошибка запроса: {e}")
        return "Ошибка геолокации"
    except Exception as e:
        logger.error(f"Непредвиденная ошибка: {e}")
        return "Ошибка обработки данных"
    
    except requests.exceptions.RequestException as e:
        logger.error(f"Ошибка запроса: {e}")
        return "Ошибка при запросе данных"
    except ValueError:
        logger.error("Ошибка декодирования JSON")
        return "Ошибка в данных"


def save_to_excel(user_data: dict, user_id: int):
    """ Сохранение данных пользователя в Excel. """
    current_date = datetime.now().strftime("%d.%m.%Y")  # ДАТА МЕСЯЦ ГОД
    
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
    
    df.to_excel(EXCEL_FILE, index=False)

async def start_handler(message: types.Message, state: FSMContext):
    """ Стартовый обработчик — запрашивает геолокацию. """
    await state.set_state(Form.location)
    
    keyboard = ReplyKeyboardBuilder()
    keyboard.add(types.KeyboardButton(text="📍 Отправить геолокацию", request_location=True))
    
    await message.answer("📍 Пожалуйста, отправьте вашу геолокацию:", reply_markup=keyboard.as_markup(resize_keyboard=True, one_time_keyboard=True))

async def process_location(message: types.Message, state: FSMContext):
    """ Обрабатывает геолокацию и определяет город. """
    if not message.location:
        await message.answer("❌ Пожалуйста, используйте кнопку для отправки геолокации.")
        return
    
    latitude = message.location.latitude
    longitude = message.location.longitude
    city = get_city_by_coordinates(latitude, longitude)
    
    await state.update_data(region=city, city=city, location=(latitude, longitude))
    await state.set_state(Form.fio)
    
    await message.answer(f"✅ Координаты получены!\n🌍 Широта: {latitude}, Долгота: {longitude}\n🏙 Определенный город: {city}\n\nТеперь введите ваше ФИО:")

async def process_fio(message: types.Message, state: FSMContext):
    """ Обработка ФИО пользователя. """
    await state.update_data(fio=message.text)
    
    builder = ReplyKeyboardBuilder()
    builder.add(types.KeyboardButton(text="Подтвердить"))
    builder.add(types.KeyboardButton(text="Изменить"))
    
    await message.answer(
        f"🔍 Проверьте ФИО: {message.text}",
        reply_markup=builder.as_markup(resize_keyboard=True, one_time_keyboard=True)
    )
    await state.set_state(Form.bs_number)

async def process_bs_number(message: types.Message, state: FSMContext):
    """ Проверка и сохранение номера базовой станции. """
    if message.text == 'Изменить':
        await state.set_state(Form.fio)
        await message.answer("Введите ФИО:")
        return
    
    text = message.text.upper().strip()
    if not re.match(r'^[A-Z]{2}-\d{6}$', text):
        await message.answer("Введите номер БС в формате VD-XXXXXX:")
        return
    
    await state.update_data(bs_number=text)
    await state.set_state(Form.stages)
    
    builder = InlineKeyboardBuilder()
    stages = ['Этап_1', 'Этап_2', 'Этап_3', 'Этап_4','Этап_5', 'Этап_6', 'Этап_7', 'Этап_8', 'Этап_9']
    for stage in stages:
        builder.button(text=f"{stage.replace('_', ' ')} ❌", callback_data=f"stage_{stage}")
    builder.button(text="✅ Сохранить", callback_data="done")
    builder.adjust(1)
    
    await message.answer("🔘 Выберите этапы (можно несколько):", reply_markup=builder.as_markup())

async def process_stages(callback: types.CallbackQuery, state: FSMContext):
    """ Выбор этапов работы. """
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
        stages = ['Этап_1', 'Этап_2', 'Этап_3', 'Этап_4','Этап_5', 'Этап_6', 'Этап_7', 'Этап_8', 'Этап_9']
        for s in stages:
            status = '✅' if s in selected else '❌'
            builder.button(text=f"{s.replace('_', ' ')} {status}", callback_data=f"stage_{s}")
        builder.button(text="✅ Сохранить", callback_data="done")
        builder.adjust(1)
        
        await callback.message.edit_reply_markup(reply_markup=builder.as_markup())
        await callback.answer()
    
    elif callback.data == 'done':
        if not selected:
            await callback.answer("Выберите хотя бы один этап!", show_alert=True)
            return
        
        user_data = await state.get_data()
        save_to_excel(user_data, callback.from_user.id)
        
        await callback.message.edit_text("✅ Данные успешно сохранены!")
        await state.clear()
        await callback.answer()

async def send_excel_handler(message: types.Message):
    """ Отправка Excel-файла с данными. """
    try:
        with open(EXCEL_FILE, 'rb') as file:
            await message.answer_document(
                types.BufferedInputFile(file.read(), filename='Данные.xlsx'),
                caption='📄 Актуальные данные'
            )
    except FileNotFoundError:
        await message.answer("Файл пуст.")

async def cancel_handler(message: types.Message, state: FSMContext):
    """ Отмена текущей сессии. """
    await state.clear()
    await message.answer('❌ Сессия завершена')

def setup_handlers(dp: Dispatcher):
    dp.message.register(start_handler, Command('start'))
    dp.message.register(send_excel_handler, Command('get_excel'))
    dp.message.register(cancel_handler, Command('cancel'))
    
    dp.message.register(process_location, Form.location)
    dp.message.register(process_fio, Form.fio)
    dp.message.register(process_bs_number, Form.bs_number)
    
    dp.callback_query.register(process_stages, Form.stages)

async def main():
    bot = Bot(token="7856193785:AAFhbj0B8TI33LuXJnwA8PimgwM07PauAg8")
    dp = Dispatcher()
    setup_handlers(dp)
    await dp.start_polling(bot)

if __name__ == '__main__':
    import asyncio
    asyncio.run(main())
