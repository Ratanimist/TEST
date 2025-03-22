from dotenv import load_dotenv
import os
from aiogram import Dispatcher, Bot
from aiogram.filters import Command
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
import asyncio
import shutil
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
import win32com.client as win32
from aiogram.types import FSInputFile


PATH: os.PathLike = None
QUESTION_ID: int = None

load_dotenv()
token = os.getenv("TOKEN")
bot = Bot(token=token)
dp = Dispatcher()

class Question(StatesGroup):
    current_question = State()

@dp.message(Command("start"))
async def start(message):
    keyboard = ReplyKeyboardMarkup(
        keyboard=[
            [
                KeyboardButton(text='НАЧАТЬ')
            ]
        ],
        resize_keyboard=True
    )

    await message.answer("Здравствуйте! Данный бот позволяет создать PDF файл. Хотите начать?", reply_markup=keyboard)

@dp.message(lambda message: message.text == "НАЧАТЬ")
async def start_button(message, state):
    global PATH

    try:
        with open('access.txt', 'r') as file:
            users = file.readlines()
    except:
        users = []
        print("Файл не найден!")

    user = message.from_user
    if user.username not in users:
        await message.answer(text="Доступ к данному боту закрыт")
    else:
        await message.answer(text='Привет!')
        dirs = os.listdir('files_excel')
        if user.username not in dirs:
            os.makedirs(os.path.join('files_excel', user.username))
            shutil.copytree(os.path.join('files_excel', 'files_start'), os.path.join('files_excel', user.username, 'files_start'))
        
        dirs_user = os.listdir(os.path.join('files_excel', user.username))
        if "files_prem" in dirs_user:
            PATH = os.path.join('files_excel', user.username, 'files_prem')
        else:
            PATH = os.path.join('files_excel', user.username, 'files_start')
        
        answer = pd.read_excel(os.path.join(PATH, 'otvet.xlsx'))
        answer[['Ответы число', 'Ответы текст']] = np.nan
        answer.to_excel(os.path.join(PATH, 'otvet.xlsx'), index=False)
        await controller_questions(message, state)
    
async def controller_questions(message, state):
    path = os.path.join(PATH, 'CSV.csv')
    questions = pd.read_csv(path, index_col='id', sep=';')
    last_index = questions.index[-1]
    question_idxs = questions.index
    await state.update_data(question_idx=0, last_idx=[], end_idx=last_index)
    await send_question(message, state)
    await state.set_state(Question.current_question)

async def send_question(message, state):
    data = await state.get_data()

    path = os.path.join(PATH, 'CSV.csv')
    df = pd.read_csv(path, sep=';', index_col='id')
    row = df.loc[data['question_idx'], :]

    if row['Вариант клавиатуры'] == 1:
        keyboard = ReplyKeyboardMarkup(
            keyboard=[
                [
                    KeyboardButton(text="ДА"),
                    KeyboardButton(text="НЕТ"),
                ],
                [
                    KeyboardButton(text="НАЗАД"),
                    KeyboardButton(text="ДАЛЕЕ"),
                ],
                [
                    KeyboardButton(text="ЗАВЕРШИТЬ ПРОГРАММУ"),
                ]
            ],
            resize_keyboard=True
        )
    elif row['Вариант клавиатуры'] == 2:
        buttons = [
                        [
                            KeyboardButton(text="НАЗАД"),
                            KeyboardButton(text="ДАЛЕЕ"),
                        ],
                        [
                            KeyboardButton(text="ЗАВЕРШИТЬ ПРОГРАММУ"),
                        ]
                ]
        if row['Варианты ответов'] is not np.nan:
            vars = row['Варианты ответов'].split(',')
            var_buttons = [[KeyboardButton(text=var)]for var in vars]
            buttons = var_buttons + buttons

        keyboard = ReplyKeyboardMarkup(
            keyboard=buttons,
            resize_keyboard=True
        )
    await state.update_data(type_keyboard=row['Вариант клавиатуры'], type_data=row['Тип данных'])
    await message.answer(text=row['Вопрос'], reply_markup=keyboard)
    

def process_excel_file():
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False  # Запуск без отображения окна Excel
    
    wb = excel.Workbooks.Open(os.path.abspath(os.path.join(PATH, "Spec.xlsx")))
    ws = wb.Sheets(1)
    
    # Нажатие на ячейку A1 для обновления связей
    ws.Range("A1").Value = ws.Range("A1").Value
    
    # Скрытие строк, где в столбце A значение 777
    last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # xlUp = -4162
    for row in range(1, last_row + 1):
        if ws.Cells(row, 1).Value == 777:
            ws.Rows(row).Hidden = True
    
    # Скрытие столбцов, где в строке 1 значение 777
    last_col = ws.Cells(1, ws.Columns.Count).End(-4159).Column  # xlToLeft = -4159
    for col in range(1, last_col + 1):
        if ws.Cells(1, col).Value == 777:
            ws.Columns(col).Hidden = True
    
    # Сохранение в PDF
    pdf_path = os.path.abspath(os.path.join(PATH, "Spec_new.pdf"))
    ws.ExportAsFixedFormat(0, pdf_path)  # 0 означает xlTypePDF
    
    wb.Close(SaveChanges=False)
    excel.Quit()
    
    return pdf_path

class ReportState(StatesGroup):
    waiting_for_filename = State()

async def exit(message, state: FSMContext):
    await message.answer("Введите название конечного отчета")
    await state.set_state(ReportState.waiting_for_filename)

@dp.message(ReportState.waiting_for_filename)
async def process_report_name(message, state: FSMContext):
    report_name = message.text.strip()
    if not report_name:
        await message.answer("Название отчета не может быть пустым. Введите снова:")
        return
    
    try:
        pdf_path = process_excel_file(report_name)  # Генерируем PDF с нужным именем
        pdf_file = FSInputFile(pdf_path)
        
        await message.answer_document(pdf_file)

        os.remove(pdf_path)  # Удаляем файл после отправки
    except Exception as e:
        await message.answer(f"Ошибка при обработке файла: {e}")

    await state.clear()

def process_excel_file(report_name):
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False  
    
    wb = excel.Workbooks.Open(os.path.abspath(os.path.join(PATH, "Spec.xlsx")))
    ws = wb.Sheets(1)

    # Нажатие на ячейку A1 для обновления связей
    ws.Range("A1").Value = ws.Range("A1").Value

    # Скрытие строк, где в столбце A значение 777
    last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  
    for row in range(1, last_row + 1):
        if ws.Cells(row, 1).Value == 777:
            ws.Rows(row).Hidden = True

    # Скрытие столбцов, где в строке 1 значение 777
    last_col = ws.Cells(1, ws.Columns.Count).End(-4159).Column  
    for col in range(1, last_col + 1):
        if ws.Cells(1, col).Value == 777:
            ws.Columns(col).Hidden = True

    # Сохранение в PDF с новым именем
    pdf_path = os.path.abspath(os.path.join(PATH, f"{report_name}.pdf"))
    ws.ExportAsFixedFormat(0, pdf_path)  

    wb.Close(SaveChanges=False)
    excel.Quit()

    return pdf_path



@dp.message(Question.current_question)
async def answer(message, state: FSMContext):
    if message.text == "ЗАВЕРШИТЬ ПРОГРАММУ":
        await exit(message, state)
        return
        
    data = await state.get_data()
    question_id = data['question_idx']
    answer = pd.read_excel(os.path.join(PATH, 'otvet.xlsx'))
    df = pd.read_csv(os.path.join(PATH, 'CSV.csv'), sep=';', index_col='id')

    if message.text == 'НАЗАД':
        question_id = data['last_idx'][-1]
        answer.loc[question_id, ['Ответы текст', 'Ответы число']] = np.nan
        answer.to_excel(os.path.join(PATH, 'otvet.xlsx'), index=False)
        last_idx = data['last_idx']
        if len(last_idx) == 1:
            last_idx = []
        else:
            last_idx.pop()
        await state.update_data(question_idx=question_id, last_idx=last_idx)
        await send_question(message, state)
        return
    if message.text == 'ДАЛЕЕ':
        end_idx = data['end_idx']
        if question_id + 1 < end_idx:
            last_idx = data['last_idx']
            last_idx.append(question_id)
            await state.update_data(question_idx=question_id + 1, last_idx=last_idx)
            await send_question(message, state)
            return
        else:
            await exit(message, state)
            return

    last_idx = data['last_idx'] 
    end_idx = data['end_idx']
    type_keyboard = data['type_keyboard']
    type_data = data['type_data']

    if type_keyboard == 1 and message.text != "ДА" and message.text != "НЕТ":
        await message.answer("Выберите ДА или НЕТ!")
        await send_question(message, state)
        return 
    
    last_idx.append(question_id)
    if type_data == 'float':
        try:
            answer.loc[question_id, 'Ответы число'] = float(message.text)
        except ValueError:
            await message.answer("Некорректный формат ввода, введите число.")
            await send_question(message, state)
            return
    else:
        answer['Ответы текст'] = answer['Ответы текст'].astype(str)
        answer.loc[question_id, 'Ответы текст'] = str(message.text)

    
    answer.to_excel(os.path.join(PATH, 'otvet.xlsx'), index=False)

    next_question_id = question_id + 1

    if message.text == "НЕТ":
        jump_to = df.loc[question_id, 'Переход к вопросу']
        if pd.notna(jump_to):
            try:
                next_question_id = int(jump_to)
            except ValueError:
                next_question_id = question_id + 1  # Если ошибка, идем дальше по порядку
        else:
            next_question_id = question_id + 1  # Если перехода нет, идем дальше

    if next_question_id < end_idx:
        await state.update_data(question_idx=next_question_id, last_idx=last_idx)
        await send_question(message, state)
    else:
        await state.clear()
        await message.answer("Вопросов больше нет!")


async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    print("Бот запущен!")
    asyncio.run(main())