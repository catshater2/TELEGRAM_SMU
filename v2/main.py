import logging
import pandas as pd
import re
import requests
from datetime import datetime, timedelta
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import StatesGroup, State
from aiogram.utils.keyboard import InlineKeyboardBuilder, ReplyKeyboardBuilder
from dotenv import load_dotenv
import os

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Константы
EXCEL_PROGRESS = 'progress_data.xlsx'
EXCEL_EMPLOYEES = 'employees.xlsx'

# Загрузка переменных окружения
load_dotenv()

class Form(StatesGroup):
    start_confirmation = State()
    confirm_region = State()
    select_foreman = State()
    select_team = State()
    input_bs_number = State()
    select_stage = State()
    confirm_start_time = State()
    stage_action = State()
    confirm_end_time = State()

class StageData:
    def __init__(self):
        self.stages = {}
    
    def add_stage(self, stage_number: int, start_time: str):
        self.stages[stage_number] = {
            'start': datetime.strptime(start_time, "%H:%M"),
            'end': None,
            'duration': None
        }
    
    def complete_stage(self, stage_number: int, end_time: str):
        end = datetime.strptime(end_time, "%H:%M")
        self.stages[stage_number]['end'] = end
        self.stages[stage_number]['duration'] = end - self.stages[stage_number]['start']

def normalize_time(time_str: str) -> str:
    """Нормализует время в формат HH:MM"""
    try:
        time_str = time_str.strip()
        if re.match(r'^\d{1,2}:\d{2}$', time_str):
            hours, minutes = time_str.split(':')
            return f"{int(hours):02d}:{int(minutes):02d}"
        return time_str
    except Exception as e:
        logger.error(f"Ошибка нормализации времени: {e}")
        return time_str

def shorten_name(full_name: str) -> str:
    parts = full_name.strip().split()
    if not parts:
        return "Неизвестный"
    last_name = parts[0]
    initials = []
    if len(parts) > 1:
        initials.append(parts[1][0] + '.')
    if len(parts) > 2:
        initials.append(parts[2][0] + '.')
    return f"{last_name} {' '.join(initials)}"

def get_region_by_coordinates(lat: float, lon: float) -> str:
    url = f"https://nominatim.openstreetmap.org/reverse?lat={lat}&lon={lon}&format=json&accept-language=ru"
    headers = {'User-Agent': 'ConstructionBot/1.0'}
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        data = response.json()
        return data.get('address', {}).get('state') or "Не определен"
    except Exception as e:
        logger.error(f"Ошибка геокодирования: {e}")
        return "Ошибка определения"

def get_foremen(region: str):
    df = pd.read_excel(EXCEL_EMPLOYEES)
    return df[
        (df['Регион'] == region) & 
        (df['Должность'].str.contains(r'^Бригадир', case=False, regex=True))
    ].reset_index(drop=True)

def get_team_members(region: str):
    df = pd.read_excel(EXCEL_EMPLOYEES)
    return df[
        (df['Регион'] == region) & 
        (~df['Должность'].str.contains(r'^Бригадир', case=False, regex=True))
    ].reset_index(drop=True)

def save_to_excel(user_data: dict):
    try:
        current_date = datetime.now().strftime("%d.%m.%Y")
        stages_data = user_data['stage_data'].stages
        
        rows = []
        for stage, times in stages_data.items():
            rows.append({
                'Дата': current_date,
                'Регион': user_data.get('region'),
                'Бригадир': user_data.get('foreman'),
                'Команда': ', '.join(user_data.get('team', [])),
                'Номер БС': user_data.get('bs_number', 'Не указан'),
                'Этап': stage,
                'Начало': times['start'].strftime("%H:%M"),
                'Окончание': times['end'].strftime("%H:%M") if times['end'] else None,
                'Длительность': str(times['duration']).split('.')[0] if times['duration'] else None
            })

        df = pd.DataFrame(rows)
        file_exists = os.path.exists(EXCEL_PROGRESS)

        if file_exists:
            existing_df = pd.read_excel(EXCEL_PROGRESS)
            df = pd.concat([existing_df, df], ignore_index=True)

        with pd.ExcelWriter(
            EXCEL_PROGRESS,
            engine='openpyxl',
            mode='a' if file_exists else 'w',
            if_sheet_exists='overlay' if file_exists else None
        ) as writer:
            df.to_excel(writer, index=False)
            
        logger.info("Данные успешно сохранены в Excel")
        return True
    except PermissionError as e:
        logger.error(f"Ошибка доступа к файлу: {e}")
        return False
    except Exception as e:
        logger.error(f"Ошибка сохранения данных: {e}")
        return False

async def start_handler(message: types.Message, state: FSMContext):
    await state.clear()
    builder = ReplyKeyboardBuilder()
    builder.row(
        types.KeyboardButton(text="✅ Да"),
        types.KeyboardButton(text="❌ Нет")
    )
    await message.answer(
        "Хотите начать новую работу?",
        reply_markup=builder.as_markup(resize_keyboard=True)
    )
    await state.set_state(Form.start_confirmation)

async def handle_start_confirmation(message: types.Message, state: FSMContext):
    if message.text == "✅ Да":
        builder = ReplyKeyboardBuilder()
        builder.add(types.KeyboardButton(
            text="📍 Отправить геолокацию",
            request_location=True
        ))
        await message.answer(
            "Отправьте вашу геолокацию:",
            reply_markup=builder.as_markup(resize_keyboard=True)
        )
        await state.set_state(Form.confirm_region)
    else:
        await message.answer("Работа отменена. Для начала новой работы используйте /start")
        await state.clear()

async def handle_location(message: types.Message, state: FSMContext):
    location = message.location
    region = get_region_by_coordinates(location.latitude, location.longitude).capitalize()
    
    foremen = get_foremen(region)
    if foremen.empty:
        await message.answer("❌ В вашем регионе нет бригадиров! Отправьте геолокацию снова или введите регион вручную.")
        await state.set_state(Form.confirm_region)
        return
    
    await state.update_data(region=region)
    
    builder = InlineKeyboardBuilder()
    builder.row(
        types.InlineKeyboardButton(text="✅ Да", callback_data="region_confirm"),
        types.InlineKeyboardButton(text="✏️ Ввести вручную", callback_data="region_edit")
    )
    await message.answer(
        f"Ваш регион: {region}\nВсё верно?",
        reply_markup=builder.as_markup()
    )

async def handle_manual_region(message: types.Message, state: FSMContext):
    region = message.text.strip().capitalize()
    await state.update_data(region=region)
    
    foremen = get_foremen(region)
    if foremen.empty:
        await message.answer("❌ В вашем регионе нет бригадиров! Пожалуйста, введите регион заново:")
        await state.set_state(Form.confirm_region)
        return
    
    builder = InlineKeyboardBuilder()
    for idx, row in foremen.iterrows():
        short_name = shorten_name(row['Сотрудник'])
        builder.add(types.InlineKeyboardButton(
            text=short_name,
            callback_data=f"foreman_{idx}"
        ))
    builder.adjust(1)
    
    await state.set_state(Form.select_foreman)
    await message.answer(
        "Выберите себя из списка бригадиров:",
        reply_markup=builder.as_markup()
    )

async def region_callback(callback: types.CallbackQuery, state: FSMContext):
    if callback.data == "region_edit":
        await callback.message.answer("Введите регион вручную:")
        await state.set_state(Form.confirm_region)
        return
    
    data = await state.get_data()
    region = data['region']
    
    foremen = get_foremen(region)
    if foremen.empty:
        await callback.message.answer("❌ В вашем регионе нет бригадиров! Введите регион заново:")
        await state.set_state(Form.confirm_region)
        return
    
    builder = InlineKeyboardBuilder()
    for idx, row in foremen.iterrows():
        short_name = shorten_name(row['Сотрудник'])
        builder.add(types.InlineKeyboardButton(
            text=short_name,
            callback_data=f"foreman_{idx}"
        ))
    builder.adjust(1)
    
    await state.set_state(Form.select_foreman)
    await callback.message.answer(
        "Выберите себя из списка бригадиров:",
        reply_markup=builder.as_markup()
    )

async def foreman_selected(callback: types.CallbackQuery, state: FSMContext):
    try:
        data = await state.get_data()
        region = data['region']
        foremen = get_foremen(region)
        
        foreman_id = int(callback.data.split("_")[1])
        if foreman_id >= len(foremen):
            await callback.answer("Неверный выбор!")
            return
            
        foreman = foremen.iloc[foreman_id]['Сотрудник']
        await state.update_data(foreman=foreman)
        await show_team_buttons(callback.message, state)
        
    except Exception as e:
        logger.error(f"Ошибка выбора бригадира: {e}")
        await callback.answer("Ошибка выбора, начните заново /start")

async def show_team_buttons(message: types.Message, state: FSMContext):
    data = await state.get_data()
    team = get_team_members(data['region'])
    
    if team.empty:
        await message.answer("✅ Команда не требуется. Переходим к этапам.")
        await state.update_data(team=[])
        await show_stage_buttons(message, state)
        return
    
    builder = InlineKeyboardBuilder()
    for idx, row in team.iterrows():
        short_name = shorten_name(row['Сотрудник'])
        builder.add(types.InlineKeyboardButton(
            text=f"❌ {short_name}",
            callback_data=f"member_{idx}"
        ))
    builder.adjust(2)
    builder.row(types.InlineKeyboardButton(
        text="✅ Завершить выбор",
        callback_data="team_done"
    ))
    
    await state.set_state(Form.select_team)
    await message.answer(
        "Выберите членов команды:",
        reply_markup=builder.as_markup()
    )

async def team_selection(callback: types.CallbackQuery, state: FSMContext):
    try:
        data = await state.get_data()
        region = data['region']
        team_df = get_team_members(region)
        selected = data.get('team', [])
        
        if callback.data == "team_done":
            if not selected:
                await callback.answer("Выберите хотя бы одного сотрудника!", show_alert=True)
                return
            
            await callback.message.answer("Введите номер БС в формате: две латинские буквы и шесть цифр (например VD123456):")
            await state.set_state(Form.input_bs_number)
            return
        
        member_id = int(callback.data.split("_")[1])
        if member_id >= len(team_df):
            await callback.answer("Неверный выбор!")
            return
            
        member = team_df.iloc[member_id]['Сотрудник']
        short_name = shorten_name(member)
        
        if member in selected:
            selected.remove(member)
            status = "❌"
        else:
            selected.append(member)
            status = "✅"
        
        await state.update_data(team=selected)
        
        builder = InlineKeyboardBuilder()
        for idx, row in team_df.iterrows():
            current_member = row['Сотрудник']
            current_short = shorten_name(current_member)
            current_status = "✅" if current_member in selected else "❌"
            builder.add(types.InlineKeyboardButton(
                text=f"{current_status} {current_short}",
                callback_data=f"member_{idx}"
            ))
        
        builder.adjust(2)
        builder.row(types.InlineKeyboardButton(
            text="✅ Завершить выбор",
            callback_data="team_done"
        ))
        
        await callback.message.edit_reply_markup(
            reply_markup=builder.as_markup()
        )
        await callback.answer()
        
    except Exception as e:
        logger.error(f"Ошибка выбора команды: {e}")
        await callback.answer("Ошибка выбора, начните заново /start")

async def handle_bs_number(message: types.Message, state: FSMContext):
    bs_number = message.text.strip().upper()
    if not re.match(r'^[A-Z]{2}\d{6}$', bs_number):
        await message.answer("❌ Неверный формат! Используйте две латинские буквы и шесть цифр (пример: VD123456)")
        return
    
    await state.update_data(bs_number=bs_number)
    await show_stage_buttons(message, state)

async def show_stage_buttons(message: types.Message, state: FSMContext):
    builder = InlineKeyboardBuilder()
    for i in range(1, 10):
        builder.add(types.InlineKeyboardButton(
            text=f"Этап {i}", 
            callback_data=f"stage_{i}"
        ))
    builder.adjust(3)
    
    await state.set_state(Form.select_stage)
    await message.answer(
        "Выберите текущий этап:",
        reply_markup=builder.as_markup()
    )

async def stage_selected(callback: types.CallbackQuery, state: FSMContext):
    stage = int(callback.data.split("_")[1])
    await state.update_data(current_stage=stage)
    
    builder = InlineKeyboardBuilder()
    builder.row(
        types.InlineKeyboardButton(
            text="✅ Подтвердить текущее время",
            callback_data="start_time_confirm"
        ),
        types.InlineKeyboardButton(
            text="✏️ Ввести время",
            callback_data="start_time_edit"
        )
    )
    
    await state.set_state(Form.confirm_start_time)
    await callback.message.answer(
        f"Укажите время начала для этапа {stage}:",
        reply_markup=builder.as_markup()
    )

async def handle_start_time(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    stage_data: StageData = data.get('stage_data', StageData())
    current_stage = data['current_stage']
    
    if callback.data == "start_time_confirm":
        current_time = datetime.now().strftime("%H:%M")
        
        # Завершаем предыдущий этап
        if stage_data.stages:
            last_stage = max(stage_data.stages.keys())
            if not stage_data.stages[last_stage]['end']:
                stage_data.complete_stage(last_stage, current_time)
        
        # Добавляем новый этап
        stage_data.add_stage(current_stage, current_time)
        
    elif callback.data == "start_time_edit":
        await callback.message.answer("Введите время начала этапа (ЧЧ:ММ):")
        return
    
    await state.update_data(stage_data=stage_data)
    await show_stage_actions(callback.message, state)
    await callback.answer()

async def handle_start_time_input(message: types.Message, state: FSMContext):
    try:
        time_input = normalize_time(message.text)
        if not re.match(r'^(0[0-9]|1[0-9]|2[0-3]):([0-5][0-9])$', time_input):
            raise ValueError
        
        data = await state.get_data()
        stage_data: StageData = data.get('stage_data', StageData())
        current_stage = data['current_stage']
        
        # Завершаем предыдущий этап
        if stage_data.stages:
            last_stage = max(stage_data.stages.keys())
            if not stage_data.stages[last_stage]['end']:
                stage_data.complete_stage(last_stage, time_input)
        
        # Добавляем новый этап
        stage_data.add_stage(current_stage, time_input)
        
        await state.update_data(stage_data=stage_data)
        await show_stage_actions(message, state)
    
    except Exception as e:
        logger.error(f"Ошибка ввода времени: {e}")
        await message.answer("❌ Неверный формат времени! Используйте ЧЧ:ММ (например: 09:30 или 18:00)")
        await state.set_state(Form.confirm_start_time)

async def show_stage_actions(message: types.Message, state: FSMContext):
    builder = InlineKeyboardBuilder()
    builder.add(types.InlineKeyboardButton(
        text="✅ Завершить работу",
        callback_data="complete_work"
    ))
    builder.add(types.InlineKeyboardButton(
        text="➡️ Следующий этап",
        callback_data="next_stage"
    ))
    builder.adjust(1)
    
    await state.set_state(Form.stage_action)
    await message.answer(
        "Выберите действие:",
        reply_markup=builder.as_markup()
    )

async def handle_stage_action(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    stage_data: StageData = data['stage_data']
    
    if callback.data == "complete_work":
        current_time = datetime.now().strftime("%H:%M")
        builder = InlineKeyboardBuilder()
        builder.row(
            types.InlineKeyboardButton(
                text="✅ Подтвердить время",
                callback_data="end_time_confirm"
            ),
            types.InlineKeyboardButton(
                text="✏️ Ввести время",
                callback_data="end_time_edit"
            )
        )
        await state.update_data(end_time=current_time)
        await callback.message.answer(
            f"Работа завершена в {current_time}",
            reply_markup=builder.as_markup()
        )
        await state.set_state(Form.confirm_end_time)
    
    elif callback.data == "next_stage":
        await show_stage_buttons(callback.message, state)
    
    await callback.answer()

async def handle_end_time_confirmation(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    stage_data: StageData = data['stage_data']
    
    if callback.data == "end_time_confirm":
        end_time = data['end_time']
    else:
        await callback.message.answer("Введите время окончания работы (ЧЧ:ММ):")
        return
    
    last_stage = max(stage_data.stages.keys())
    stage_data.complete_stage(last_stage, end_time)
    
    if save_to_excel(data):
        await callback.message.answer("✅ Данные о работе успешно сохранены!")
    else:
        await callback.message.answer("❌ Ошибка сохранения данных!")
    
    await state.clear()

async def handle_end_time_input(message: types.Message, state: FSMContext):
    try:
        time_input = normalize_time(message.text)
        if not re.match(r'^(0[0-9]|1[0-9]|2[0-3]):([0-5][0-9])$', time_input):
            raise ValueError
        
        data = await state.get_data()
        stage_data: StageData = data['stage_data']
        last_stage = max(stage_data.stages.keys())
        stage_data.complete_stage(last_stage, time_input)
        
        if save_to_excel(data):
            await message.answer("✅ Данные о работе успешно сохранены!")
        else:
            await message.answer("❌ Ошибка сохранения данных!")
        
        await state.clear()
    
    except Exception as e:
        logger.error(f"Ошибка ввода времени: {e}")
        await message.answer("❌ Неверный формат времени! Используйте ЧЧ:ММ (например: 09:30 или 18:00)")
        await state.set_state(Form.confirm_end_time)

async def cancel_handler(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(
        '❌ Сессия завершена. Для начала новой работы используйте /start',
        reply_markup=types.ReplyKeyboardRemove()
    )

def setup_handlers(dp: Dispatcher):
    dp.message.register(start_handler, Command('start'))
    dp.message.register(cancel_handler, Command('cancel'))
    dp.message.register(handle_start_confirmation, Form.start_confirmation)
    dp.message.register(handle_location, F.location, Form.confirm_region)
    dp.message.register(handle_manual_region, Form.confirm_region)
    dp.message.register(handle_bs_number, Form.input_bs_number)
    dp.message.register(handle_start_time_input, Form.confirm_start_time)
    dp.message.register(handle_end_time_input, Form.confirm_end_time)
    
    dp.callback_query.register(region_callback, F.data.startswith("region_"))
    dp.callback_query.register(foreman_selected, F.data.startswith("foreman_"))
    dp.callback_query.register(team_selection, Form.select_team)
    dp.callback_query.register(stage_selected, F.data.startswith("stage_"))
    dp.callback_query.register(handle_start_time, Form.confirm_start_time)
    dp.callback_query.register(handle_stage_action, Form.stage_action)
    dp.callback_query.register(handle_end_time_confirmation, Form.confirm_end_time)

async def main():
    bot = Bot(token=os.getenv("BOT_TOKEN"))
    dp = Dispatcher()
    setup_handlers(dp)
    await dp.start_polling(bot)

if __name__ == '__main__':
    import asyncio
    asyncio.run(main())
