import logging
import pandas as pd
import re
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import StatesGroup, State
from aiogram.utils.keyboard import InlineKeyboardBuilder, ReplyKeyboardBuilder

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

EXCEL_FILE = 'user_data.xlsx'

class Form(StatesGroup):
    region = State()
    fio = State()
    bs_number = State()
    stages = State()

def save_to_excel(user_data: dict, user_id: int):
    new_data = pd.DataFrame([{
        'user_id': user_id,
        'Регион': user_data.get('region'),
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
    await state.set_state(Form.region)
    await message.answer("Введите регион:")

async def process_region(message: types.Message, state: FSMContext):
    await state.update_data(region=message.text)
    await state.set_state(Form.fio)
    await message.answer("Введите ФИО:")

async def process_fio(message: types.Message, state: FSMContext):
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
        builder.button(
            text=f"{stage.replace('_', ' ')} ❌", 
            callback_data=f"stage_{stage}"
        )
    builder.button(text="✅ Сохранить", callback_data="done")
    builder.adjust(1)
    
    await message.answer(
        "🔘 Выберите этапы (можно несколько):",
        reply_markup=builder.as_markup()
    )

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
        stages = ['Этап_1', 'Этап_2', 'Этап_3', 'Этап_4','Этап_5', 'Этап_6', 'Этап_7', 'Этап_8', 'Этап_9']
        for s in stages:
            status = '✅' if s in selected else '❌'
            builder.button(
                text=f"{s.replace('_', ' ')} {status}", 
                callback_data=f"stage_{s}"
            )
        builder.button(text="✅ Сохранить", callback_data="done")
        builder.adjust(1)
        
        await callback.message.edit_reply_markup(
            reply_markup=builder.as_markup()
        )
        await callback.answer()
    
    elif callback.data == 'done':
        if not selected:
            await callback.answer("Выберите хотя бы один этап!", show_alert=True)
            return
        
        user_data = await state.get_data()
        save_to_excel(user_data, callback.from_user.id)
        await callback.message.edit_text("Данные успешно сохранены!")
        await state.clear()
        await callback.answer()

async def send_excel_handler(message: types.Message):
    try:
        with open(EXCEL_FILE, 'rb') as file:
            await message.answer_document(
                types.BufferedInputFile(file.read(), filename='Данные.xlsx'),
                caption='Актуальные данные'
            )
    except FileNotFoundError:
        await message.answer("Файлпуст")

async def cancel_handler(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer('Сессия завершена')

def setup_handlers(dp: Dispatcher):
    dp.message.register(start_handler, Command('start'))
    dp.message.register(send_excel_handler, Command('get_excel'))
    dp.message.register(cancel_handler, Command('cancel'))
    
    dp.message.register(process_region, Form.region)
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