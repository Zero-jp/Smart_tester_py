import datetime

import pyautogui
from win32api import GetSystemMetrics
import time
import pytesseract # оптический анализатор(нейронка)
import pandas as pd # считывание из таблицы входных данных
import math
from PIL import Image as PilImage, ImageFilter
from langdetect import detect
import pyperclip
import sys
from progress.bar import IncrementalBar
from pynput import keyboard
from tkinter import *
from tkinter import ttk
from threading import Thread
cmb = [{keyboard.Key.esc}]
current = set()
def execute():
    print("Detected hotkey")

def on_press(key):
    if any([key in z for z in cmb]):
        current.add(key)
        if any(all(k in current for k in z) for z in cmb):
            execute()

def on_release(key):
    if any([key in z for z in cmb]):
        raise SystemExit(1)

pyautogui.FAILSAFE = False
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
user_path_IGOR = "SmartTesterForBoas"
user_path_LERA = "Smart_tester_py"
Column_Coords = [0, 0]
One_C_Coords = [280, 70]
Exit_Coords_Little = [0, 0]
Exit_Coords_Big = [0, 0]
# C:\Users\Valeria\Desktop\Smart_tester_py
# C:\Users\drtar\Desktop\SmartTesterForBoas
# C:\Users\Xeon\Desktop\Smart_tester_py
path = r'C:\Users\Valeria\Desktop\Smart_tester_py'


def lookOnScreen(lang: str, is_obrez: bool):#, is_centering: False):
    global path
    temp_screen = pyautogui.screenshot(path+r'\images\temp.bmp')
    img_seryy = temp_screen.convert("L")# .save(r'C:\Users\drtar\Desktop\SmartTesterForBoas\programm\temp_kraya.bmp')
    img_seryy.filter(ImageFilter.DETAIL).save(path+r"\images\temp_kraya.bmp")
    link_obrez = path+r"\images\temp_kraya_obrez.bmp"
    link = path+r"\images\temp_kraya.bmp"
    l = ""
    # if is_centering:
    #     temp_kraya = PilImage.open(link).crop((
    #         (GetSystemMetrics(0) - 700) // 2, # GetSystemMetrics(0) - width
    #         (GetSystemMetrics(1) - 300) // 2,
    #         (GetSystemMetrics(0) + 700) // 2,
    #         (GetSystemMetrics(1) + 300) // 2)
    #     )
    #     temp_kraya.save(path+r"\images\problem_wind.bmp")
    #     return r"\images\problem_wind.bmp"
    if is_obrez:
        temp_kraya = PilImage.open(link)
        im_crop = temp_kraya.crop((Column_Coords[0], 0, Column_Coords[1], GetSystemMetrics(1)))
        im_crop.save(link_obrez)
        l = link_obrez
    else:
        l = link
    data = pytesseract.image_to_data(l, lang=lang, config='--oem 3 --psm 12 words', output_type=pytesseract.Output.DICT)

    return data

def centerWordSearch(word: str, lang: str, cut: bool, screen_shot = 'nothing'):
    # поиск слова на экране
    # получить все данные из изображения
    data = {}
    if screen_shot == 'nothing':
        data = lookOnScreen(lang, cut)#, False)
    else:
        data = pytesseract.image_to_data(screen_shot, lang=lang, config='--oem 3 --psm 12', output_type=pytesseract.Output.DICT)
    # получение координат
    x = y = 0.1
    for d in data["text"]:
        if d.find(word) != -1:
            word_id = data["text"].index(d)
            x = data["left"][word_id] + data["width"][word_id] / 2
            y = data["top"][word_id] + data["height"][word_id] / 2
    if cut:
        return [int(x)+Column_Coords[0], int(y)]
    else:
        return [int(x), int(y)-7] # 7 - чтоб курсором перекрыть текст


def dist(p1, p2):
    return math.hypot(p2[0] - p1[0], p2[1] - p1[1])


def neerestFigure(link: str, coords: list):
    # получение координат
    minDist = 1000
    imageCoords = [0, 0]
    for pos in pyautogui.locateAllOnScreen(link):
        # поиск ближайшей точки
        d = dist(pos, coords)
        if d < minDist and coords[0]<pos[0]:
            imageCoords = [pos[0], pos[1]]
            minDist = d
    #print(imageCoords)
    return imageCoords

def neerestWord(word_main: str, word_sub: str, lang: str):
    data = lookOnScreen(lang, False)#, False)
    main_coords = []
    sub_coords = []
    meanCoords = [0, 0]
    for i in range(len(data["text"])):
        if data["text"][i].find(word_main) != -1 and data["text"][i+1].find(word_sub) != -1:
            main_coords = [data["left"][i] + data["width"][i], data["top"][i] + data["height"][i] / 2]
            sub_coords = [data["left"][i+1] + data["width"][i+1], data["top"][i+1] + data["height"][i+1] / 2]
            global Column_Coords
            Column_Coords = [data["left"][i]-1, data["left"][i+1] + data["width"][i+1]]
    meanCoords = [(main_coords[0] + sub_coords[0]) / 2, (main_coords[1] + sub_coords[1]) / 2]
    return meanCoords

def findWordWithPicrure(skrin_location: str, word: str, lang: str):
    screenshot_pos = pyautogui.locateOnScreen(skrin_location)
    word_center = centerWordSearch(word, lang, False, skrin_location)
    return [word_center[0]+screenshot_pos.left, word_center[1]+screenshot_pos.top]

def WaitingUntilFind(sign_word: str) -> bool:
    screenShot_data = lookOnScreen('rus', False)#, False)
    for i in range(len(screenShot_data["text"])):
        if screenShot_data["text"][i].find(sign_word) != -1:
            return True
    return False


# Забираем данные для прохода по полям

data = pd.DataFrame(pd.read_excel(
    io=path+r'\Копия (Агрегатор Т1) ТК для тестирования интеграции между ELMA и IMOS (3).xlsx',
    engine='openpyxl'))
out_tab = data.copy()
frame = path+r'\images\frame.bmp'
full_frame = path+r'\images\full_frame.bmp'
ok_button = path+r'\images\ok_button.bmp'
error_wind = path+r'\images\error_wind.bmp'
active_ok_button = path+r'\images\active_ok_button.bmp'
active_ok_button_error = path+r'\images\active_ok_button_error.bmp'
text_field = path+r'\images\text_field.bmp'
use_config_wind = path+r'\images\use_config_wind.bmp'
active_yes_button = path+r'\images\active_yes_button.bmp'
show_spec_button = path+r'\images\show_spec_button.bmp'
one_c_button = path+r'\images\one_c_button.jpg'
exit_button = path+r'\images\exit_button.bmp'
config_set_button = path+r'\images\config_set_button.bmp'
ok_cancel_toolbar = path+r'\images\ok_cancel_toolbar.bmp'
error_wind_color = path+r'\images\error_wind_color.bmp'
window_color_1 = path+r'\images\window_color_1.bmp'
window_color_2 = path+r'\images\window_color_2.bmp'
window_color_3 = path+r'\images\window_color_3.bmp'
error_logo = path+r'\images\error_logo.bmp'
start_name_project = path+r'\images\start_name_project.jpg'

colum_name = []
columns_count = data.shape[1]
Temp_Data = [0, 0]
rows_count = data.shape[0]
for index in range(columns_count):
    colum_name.append(data.columns[index].split(' '))
out_tab["Статус"] = ["False"] * rows_count
out_tab["Ссылка на json"] = ["Отсутсвует"] * rows_count
print(colum_name[0][0])
print(colum_name[0][1])
print(colum_name[1][0])
print(colum_name[1][1])

sleep_timer = .1
running = False
exit = False

def startTesting(index, row):
    global path
    global data
    global out_tab
    global frame
    global full_frame
    global ok_button
    global error_wind
    global active_ok_button
    global active_ok_button_error
    global text_field
    global use_config_wind
    global active_yes_button
    global show_spec_button
    global one_c_button
    global exit_button
    global config_set_button
    global ok_cancel_toolbar
    global error_wind_color
    global window_color_1
    global window_color_2
    global window_color_3
    global error_logo
    global start_name_project
    global colum_name
    global columns_count
    global Temp_Data
    global rows_count

    global Exit_Coords_Little
    global Exit_Coords_Big

    mylist = [1, 2, 3, 4, 5, 6, 7]
    try:
        while(WaitingUntilFind("фасадов") == False):
            pyautogui.sleep(sleep_timer)
        pyautogui.sleep(sleep_timer)
        bar = IncrementalBar("Тест-кейс " + str(index+1), max=len(mylist))
        # Нахождение названия свойства и элемента со списком
        first_name = neerestWord(colum_name[0][0], colum_name[0][1], 'rus')
        pyautogui.leftClick(neerestFigure(frame, first_name))
        pyautogui.sleep(.2)
        print(detect(row[0]))
        if detect(row[0]) == 'ru':
            pyautogui.leftClick(centerWordSearch(row[0], 'rus', True))
        else:
            pyautogui.leftClick(centerWordSearch(row[0], 'eng', True))
        pyautogui.sleep(.2)
        second_name = neerestWord(colum_name[1][0], colum_name[1][1], 'rus')
        pyautogui.leftClick(neerestFigure(frame, second_name))
        # Нахождение нужного значения
        for index_value in range(columns_count-1):
            index_value += 1
            if len(row[index_value].split(" ")) == 1:
                if detect(row[index_value]) == 'ru' or detect(row[index_value]) == 'uk' or detect(row[index_value]) == 'bk' \
                        or detect(row[index_value]) == 'bg' or detect(row[index_value]) == 'mk':
                    pyautogui.leftClick(centerWordSearch(row[index_value], 'rus', True))
                else:
                    pyautogui.leftClick(centerWordSearch(row[index_value], 'eng', True))
            elif len(row[index_value].split(" ")) == 2:
                words = row[index_value].split(" ")
                if detect(row[index_value]) == 'ru' or detect(row[index_value]) == 'uk' or detect(row[index_value]) == 'bk' \
                        or detect(row[index_value]) == 'bg' or detect(row[index_value]) == 'mk':
                    pyautogui.leftClick(neerestWord(words[0], words[1], 'rus'))
                else:
                    pyautogui.leftClick(neerestWord(words[0], words[1], 'eng'))
            bar.next()
            pyautogui.sleep(.2)
            pyautogui.press('tab')
            pyautogui.sleep(.1)
            pyautogui.press('tab')
            pyautogui.sleep(.1)
            pyautogui.press('down')
            pyautogui.sleep(.1)
            pyautogui.press('tab')
            pyautogui.sleep(.1)
            pyautogui.press('down')
            pyautogui.sleep(.1)
            # Переход на вкладку Дополнительные параметры
            pyautogui.leftClick(neerestWord("Дополнительные", "параметры", 'rus'))
            while (WaitingUntilFind("ФИО") == False):
                pyautogui.sleep(sleep_timer)
            if index == 0:
                Temp_Data = centerWordSearch("ELMA", 'eng', False)
            pyautogui.doubleClick(Temp_Data)
            pyautogui.keyDown('ctrl')
            pyautogui.keyDown('a')
            pyautogui.keyDown('backspace')
            pyautogui.keyUp('backspace')
            pyautogui.keyUp('a')
            pyautogui.keyUp('ctrl')
            pyautogui.write(str(index+1)) #2023-03-07|2023-03-0714:04:06.055323
            pyautogui.leftClick(pyautogui.locateCenterOnScreen(active_ok_button)) # centerWordSearch("OK", 'eng', False)
            bar.next()
            pyautogui.sleep(1)
            pyautogui.press('enter')
            pyautogui.sleep(.2)
            pyautogui.leftClick(pyautogui.locateCenterOnScreen(show_spec_button))
            bar.next()
            while (WaitingUntilFind("Спецификация") == False):
                pyautogui.sleep(sleep_timer)
            pyautogui.leftClick(One_C_Coords)
            bar.next()
            while (WaitingUntilFind("успешно") == False):
                pyautogui.sleep(sleep_timer)
            bar.next()
            if index == 0:
                Exit_Coords_Little = pyautogui.locateOnScreen(exit_button, region=(
                    (GetSystemMetrics(0) - 700)// 2,
                    (GetSystemMetrics(1) - 300) // 2,
                    700, 300), confidence=0.95)
                # Exit_Coords_Little = pyautogui.locateCenterOnScreen(exit_button)
            pyautogui.leftClick(Exit_Coords_Little)
            pyautogui.sleep(.2)
            bar.next()
            if index == 0:
                Exit_Coords_Big = pyautogui.locateCenterOnScreen(exit_button)
            pyautogui.leftClick(Exit_Coords_Big)
            pyautogui.sleep(.2)
            out_tab["Статус"][index] = "Succes"
            out_tab["Ссылка на json"][index] = pyperclip.paste()
            bar.next()
            pyautogui.leftClick(pyautogui.locateCenterOnScreen(config_set_button))
            bar.finish()
        # create excel writer
        writer = pd.ExcelWriter(path+r'\logs.xlsx')# посмотретт формат csv
        # write dataframe to excel sheet named 'marks'
        out_tab.to_excel(writer, 'marks')
        # save the excel file
        writer.save()
    except Exception:
        if centerWordSearch("Спецификация", 'rus', False) != [0,-7]:
            print("Спец ошибка")
            print(centerWordSearch("Спецификация", 'rus', False))
            pyautogui.leftClick(pyautogui.locateCenterOnScreen(exit_button))
        if centerWordSearch("Буфер", 'rus', False) != [0,-7]:
            print("Буф ошибка")
            pyautogui.leftClick(pyautogui.locateCenterOnScreen(exit_button))
            pyautogui.leftClick(pyautogui.locateCenterOnScreen(exit_button))
        pyautogui.leftClick(pyautogui.locateCenterOnScreen(config_set_button))
        e = sys.exc_info()[1]
        out_tab["Статус"][index] = e.args[0]

def primaryStart(pb):
    pb.start()
    global running
    running = True

def primaryPause(pb):
    global running
    if running == True:
        pb.stop()
        running = False
    else:
        pb.start()
        running = True

inProcess = False
def primaryExit(pb):
    pb.stop()
    global exit
    print("Попался мелкий выход")
    if inProcess == False:
        window.destroy()
        raise SystemExit(1)
    if exit == False:
        exit = True

readyToExit = False
def fatalExit(window):
    # global readyToExit
    # while readyToExit != True :
    #     pyautogui.sleep(sleep_timer)
    print("Попался капитальный выход")
    window.destroy()
    sys.exit(1)

def doSomethingOnExit():
    while exit == False:
        print("пауза")
        time.sleep(2)


def window():
    window = Tk()
    window.title("Бот-тестировщик")
    # window.geometry('400x250')
    # window.resizable(False, False)
    labelStatus = Label(window, text="Статус:")
    labelStatus.grid(column=0, row=0)
    pb = ttk.Progressbar(window, orient="horizontal", length=150, mode="indeterminate")
    pb.grid(column=1, row=0)
    labelTotalCount = Label(window, text="Всего тест-кейсов:")
    labelTotalCount.grid(column=2, row=0)
    labelCount = Label(window, text=rows_count)
    labelCount.grid(column=3, row=0)
    btnStart = Button(window, text="Старт", bg="#29f716", activebackground="#82f078", command=lambda:  primaryStart(pb))
    btnStart.grid(column=0, row=1)
    btnPause = Button(window, text="Пауза", bg="#2279f2", activebackground="#78aaf0", command=lambda: primaryPause(pb))
    btnPause.grid(column=1, row=1)
    btnStop = Button(window, text="Стоп", bg="#f22222", activebackground="#f05959", command= lambda: primaryExit(pb))
    btnStop.grid(column=2, row=1)
    # window.resizable(True, True)
    # window.update_idletasks()
    # window.overrideredirect(True)
    # window.bind('<ESC>', primaryExit(pb))
    # window.anchor("center")
    # window.protocol("WM_DELETE_WINDOW", primaryExit(pb))
    window.mainloop()

while exit == False:
    t = Thread(target=window, daemon=True)
    t.start()

    for index, row in data.iterrows():
        inProcess = True
        # print("Start")
        if running == False:
            while running == False:
                if (exit == True):
                    raise SystemExit(1)
                # print("пауза")
                # print(exit)
                # print(running)
                # ожидаем повторного нажатия кнопочки
                time.sleep(2)
            # print("Насквозь")
        if (index != 0):
            print("Пройден тест-кейс "+str(index))
        startTesting(index, row)
        time.sleep(0.2)
        if exit == True:
            break
            #t.join()
            #readyToExit = True
    raise SystemExit(1)