import datetime

import pyautogui
from win32api import GetSystemMetrics
import time
import pytesseract # оптический анализатор(нейронка)
import pandas as pd # считывание из таблицы входных данных
import math
import statistics
from PIL import Image, ImageFilter
from langdetect import detect
import pyperclip
import sys
import pathlib
from pathlib import Path
from progress.bar import IncrementalBar

pyautogui.FAILSAFE = False
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
user_path_IGOR = "SmartTesterForBoas"
user_path_LERA = "Smart_tester_py"
Column_Coords = [0, 0]
# Window_Coords = [0, 0, 0, 0]

# def leftClick(coord):
#     # клик
#     win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, coord[0], coord[1])
#     # time.sleep(.1)
#     win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, coord[0], coord[1])
#     print("Click.")
def lookOnScreen(lang: str, is_obrez: bool):
    # поиск слова на экране
    # 1018 294  1205 536
    temp_screen = pyautogui.screenshot(r'..\images\temp.bmp')
    img_seryy = temp_screen.convert("L")# .save(r'C:\Users\drtar\Desktop\SmartTesterForBoas\programm\temp_kraya.bmp')
    # temp_kraya = img_seryy.filter(ImageFilter.FIND_EDGES)#.save(r'C:\Users\drtar\Desktop\SmartTesterForBoas\programm\temp_kraya.bmp')
    img_seryy.filter(ImageFilter.DETAIL).save(r"..\images\temp_kraya.bmp")
    # temp_kraya = img_seryy.filter(ImageFilter.FIND_EDGES).save(r'C:\Users\drtar\Desktop\SmartTesterForBoas\programm\temp_kraya.bmp')
    # temp_kraya = img_seryy.filter(ImageFilter.SMOOTH)
    # temp_kraya = temp_kraya.filter(ImageFilter.EMBOSS).save(r'C:\Users\drtar\Desktop\SmartTesterForBoas\programm\temp_kraya.bmp')
    # temp_kraya = img_seryy.filter(ImageFilter.SHARPEN)
    # temp_kraya = temp_kraya.filter(ImageFilter.EMBOSS).save(r'C:\Users\drtar\Desktop\SmartTesterForBoas\programm\temp_kraya.bmp')
    # file_path = r'C:\Users\drtar\Desktop\SmartTesterForBoas\programm\temp_kraya.bmp'
    # with Image.open(file_path) as img:
    #     img.load()
    #     for _ in range(3):
    #         img = img.filter(ImageFilter.MaxFilter(3))
    #     img.save(r'C:\Users\drtar\Desktop\SmartTesterForBoas\programm\temp_kraya.bmp')
    # получить все данные из изображения
    # temp_kraya = temp_kraya.filter(ImageFilter.MaxFilter(1)).save(
    #    r'C:\Users\drtar\Desktop\SmartTesterForBoas\programm\temp_kraya.bmp')
    #print(lang)
    link_obrez = r"..\images\temp_kraya_obrez.bmp"
    link = r"..\images\temp_kraya.bmp"
    l = ""
    if is_obrez:
        #print(Column_Coords)
        temp_kraya = Image.open(link)
        im_crop = temp_kraya.crop((Column_Coords[0], 0, Column_Coords[1], GetSystemMetrics(1)))
        im_crop.save(link_obrez)
        l = link_obrez
    # elif (Window_Coords[0] != 0 and Window_Coords[1] != 0 and Window_Coords[2] != 0 and Window_Coords[3] != 0):
    #     temp_kraya = Image.open(link)
    #     im_crop = temp_kraya.crop((Window_Coords[0], Window_Coords[1], Window_Coords[2], Window_Coords[3]))
    #     im_crop.save(link)
    #     l = link
    else:
        l = link
    # print(Window_Coords)
    if lang != 'eng':
        data = pytesseract.image_to_data(l, lang=lang, config='--oem 3 --psm 12', output_type=pytesseract.Output.DICT)
    else:
        data = pytesseract.image_to_data(l, lang=lang, config='--oem 3 --psm 12',
                                         output_type=pytesseract.Output.DICT)
    print(data)

    return data

# def listWordSearch():
#     temp_screen = pyautogui.screenshot('temp.bmp')
#     img_seryy = temp_screen.convert("L")
#     list_star = pyautogui.locateCenterOnScreen(r'C:\Users\drtar\Desktop\SmartTesterForBoas\images\list_start.bmp')
#     print(list_star)
#     pyautogui.leftClick(list_star)


def centerWordSearch(word: str, lang: str, cut: bool, screen_shot = 'nothing'):
    # поиск слова на экране
    # получить все данные из изображения
    data = {}
    if screen_shot == 'nothing':
        data = lookOnScreen(lang, cut)
    else:
        data = pytesseract.image_to_data(screen_shot, lang=lang, config='--oem 3 --psm 12', output_type=pytesseract.Output.DICT)
        #print(data)
    # получение координат
    #print("искомое слово:" + word)
    #print(data["text"])
    x = y = 0.1
    for d in data["text"]:
        if d.find(word) != -1:
            #print(word)
            #print(d)
            word_id = data["text"].index(d)
            x = data["left"][word_id] + data["width"][word_id] / 2
            y = data["top"][word_id] + data["height"][word_id] / 2
            #print("["+str(x)+","+str(y)+"]")
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
    # print(word_main)
    data = lookOnScreen(lang, False)
    main_coords = []
    sub_coords = []
    meanCoords = [0, 0]
    for i in range(len(data["text"])):
        # print(d)

        if data["text"][i].find(word_main) != -1 and data["text"][i+1].find(word_sub) != -1:
            # print("Main!")
            main_coords = [data["left"][i] + data["width"][i], data["top"][i] + data["height"][i] / 2]
            sub_coords = [data["left"][i+1] + data["width"][i+1], data["top"][i+1] + data["height"][i+1] / 2]
            global Column_Coords
            Column_Coords = [data["left"][i]-1, data["left"][i+1] + data["width"][i+1]]
    meanCoords = [(main_coords[0] + sub_coords[0]) / 2, (main_coords[1] + sub_coords[1]) / 2]
    # print(meanCoords)
    return meanCoords
# def retriveSettingWindCord():
#     word_main = "Параметры"
#     word_sub = "заказа"
#     data = lookOnScreen("rus", False)
#     left_top_coords = []
#     right_down_coords = []
#     global Window_Coords
#     for i in range(len(data["text"])):
#         # print(d)
#
#         if data["text"][i].find(word_main) != -1 and data["text"][i+1].find(word_sub) != -1:
#             # print("Main!")
#             Window_Coords[0], Window_Coords[1] = data["left"][i], data["top"][i]
#             Window_Coords[2], Window_Coords[3] = GetSystemMetrics(0) - data["left"][i], GetSystemMetrics(2) - data["top"][i]
#     #Window_Coords = [left_top_coords[0] , right_down_coords]

# 2023-03-07|2023-03-0714:04:06.055323
def findWordWithPicrure(skrin_location: str, word: str, lang: str):
    screenshot_pos = pyautogui.locateOnScreen(skrin_location)
    word_center = centerWordSearch(word, lang, False, skrin_location)
    #print("Поиск со скриншотом. Отладка:")
    #print(word_center)
    #print(screenshot_pos)
    #print([word_center[0]+screenshot_pos.left, word_center[1]+screenshot_pos.top])
    return [word_center[0]+screenshot_pos.left, word_center[1]+screenshot_pos.top]


# Забираем данные для прохода по полям
data = pd.DataFrame(pd.read_excel(
    io="..\Копия (Агрегатор Т1) ТК для тестирования интеграции между ELMA и IMOS (3).xlsx",
    engine='openpyxl'))
out_tab = data.copy()
frame = r'..\images\frame.bmp'
full_frame = r'..\images\full_frame.bmp'
ok_button = r'..\images\ok_button.bmp'
error_wind = r'..\images\error_wind.bmp'
active_ok_button = r'..\images\active_ok_button.bmp'
active_ok_button_error = r'..\images\active_ok_button_error.bmp'
text_field = r'..\images\text_field.bmp'
use_config_wind = r'..\images\use_config_wind.bmp'
active_yes_button = r'..\images\active_yes_button.bmp'
show_spec_button = r'..\images\show_spec_button.bmp'
one_c_button = r'..\images\one_c_button.jpg'
exit_button = r'..\images\exit_button.bmp'
config_set_button = r'..\images\config_set_button.bmp'
ok_cancel_toolbar = r'..\images\ok_cancel_toolbar.bmp'
error_wind_color = r'..\images\error_wind_color.bmp'
window_color_1 = r'..\images\window_color_1.bmp'
window_color_2 = r'..\images\window_color_2.bmp'
window_color_3 = r'..\images\window_color_3.bmp'
error_logo = r'..\images\error_logo.bmp'
start_name_project = r'..\images\start_name_project.jpg'
# retriveSettingWindCord()

# print(data.columns[0])
colum_name = []
columns_count = data.shape[1]
Temp_Data = [0, 0]
#print(data.shape[0])
rows_count = data.shape[0]
for index in range(columns_count):
    colum_name.append(data.columns[index].split(' '))
out_tab["Статус"] = ["False"]*rows_count
out_tab["Ссылка на json"] = ["Отсутсвует"]*rows_count
# print(colum_name[0][0])
# print(colum_name[0][1])
#pyautogui.leftClick(neerestWord(colum_name[0][0], colum_name[0][1]))

mylist = [1, 2, 3, 4, 5, 6, 7]
for index, row in data.iterrows():
    bar = IncrementalBar("Тест-кейс " + str(index+1), max=len(mylist))
    # try:
    # Нахождение названия свойства и элемента со списком
    # print(row[0] + '|' + row[1])
    first_name = neerestWord(colum_name[0][0], colum_name[0][1], 'rus')
    pyautogui.leftClick(neerestFigure(frame, first_name))
    pyautogui.sleep(.2)
    # listWordSearch()
    # print(row[0])
    # print(detect(row[0]))
    if detect(row[0]) == 'ru':
        pyautogui.leftClick(centerWordSearch(row[0], 'rus', True))
    # elif detect(row[0]) == 'en' or detect(row[0]) == 'cy':
    else:
        pyautogui.leftClick(centerWordSearch(row[0], 'eng', True))
    pyautogui.sleep(.2)
    #print(colum_name[1][0] + '|' + colum_name[1][1])
    second_name = neerestWord(colum_name[1][0], colum_name[1][1], 'rus')
    pyautogui.leftClick(neerestFigure(frame, second_name))
    #print("ОТЛАДКА!!!!")
    # Нахождение нужного значения
    for index_value in range(columns_count-1):
        try:
            index_value += 1
            #print(row[index_value] + detect(row[index_value]))
            if len(row[index_value].split(" ")) == 1:
                if detect(row[index_value]) == 'ru' or detect(row[index_value]) == 'uk' or detect(row[index_value]) == 'bk' \
                        or detect(row[index_value]) == 'bg' or detect(row[index_value]) == 'mk':
                    pyautogui.leftClick(centerWordSearch(row[index_value], 'rus', True))
                # elif detect(row[0]) == 'en' or detect(row[0]) == 'cy':
                else:
                    pyautogui.leftClick(centerWordSearch(row[index_value], 'eng', True))
            elif len(row[index_value].split(" ")) == 2:
                words = row[index_value].split(" ")
                if detect(row[index_value]) == 'ru' or detect(row[index_value]) == 'uk' or detect(row[index_value]) == 'bk' \
                        or detect(row[index_value]) == 'bg' or detect(row[index_value]) == 'mk':
                    pyautogui.leftClick(neerestWord(words[0], words[1], 'rus'))
                # elif detect(row[0]) == 'en' or detect(row[0]) == 'cy':
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
            pyautogui.leftClick(neerestWord("Дополнительные", "параметры", 'rus'))
            # print("ОТЛАДКА!!!!")
            # global Temp_Data

            pyautogui.sleep(.4)
            # deal_count_cords = centerWordSearch(str(index_value), 'rus', False)
            if index_value == 1:
                Temp_Data = centerWordSearch("NumberOfSave", 'eng', False)
            pyautogui.doubleClick(Temp_Data)
            pyautogui.keyDown('ctrl')
            pyautogui.keyDown('a')
            pyautogui.keyDown('backspace')
            pyautogui.keyUp('backspace')
            pyautogui.keyUp('a')
            pyautogui.keyUp('ctrl')
            pyautogui.write(str(index_value+1)) #2023-03-07|2023-03-0714:04:06.055323
            pyautogui.leftClick(centerWordSearch("OK", 'eng', False))
            bar.next()
            pyautogui.sleep(1)
            # if pyautogui.locateCenterOnScreen(use_config_wind) != None:
            # pyautogui.leftClick(pyautogui.locateCenterOnScreen(active_yes_button))# findWordWithPicrure(use_config_wind, "Да", 'rus'))
            pyautogui.press('enter')
            pyautogui.sleep(.2)
            # if pyautogui.locateCenterOnScreen(error_wind_color) != None:
            #     pyautogui.leftClick(centerWordSearch("OK", 'eng', False))# findWordWithPicrure(use_config_wind, "Да", 'rus'))
            #     pyautogui.sleep(1)
            #     Column_Coords[0] = pyautogui.locateOnScreen(window_color_1)['left']
            #     Column_Coords[1] = pyautogui.locateOnScreen(window_color_1)['top']
            #     #print("Отладочка:")
            #     color_cord = centerWordSearch("Белый", 'rus', False, window_color_1)
            #     pyautogui.sleep(.2)
            #     #print(color_cord)
            #     if color_cord == [0, -7]:
            #         Column_Coords[0] = pyautogui.locateOnScreen(window_color_2)['left']
            #         Column_Coords[1] = pyautogui.locateOnScreen(window_color_2)['top']
            #         color_cord = centerWordSearch("Графит", 'rus', False, window_color_2)
            #         pyautogui.sleep(.2)
            #     pyautogui.leftClick(centerWordSearch("OK", 'eng', False))
            #     pyautogui.sleep(1)
            #     if pyautogui.locateCenterOnScreen(error_wind_color) != None:
            #         pyautogui.leftClick(color_cord)
            #         pyautogui.sleep(.2)
            #         Column_Coords[0] = pyautogui.locateOnScreen(window_color_3)['left']
            #         Column_Coords[1] = pyautogui.locateOnScreen(window_color_3)['top']
            #         color_cord = centerWordSearch("Мателюкс", 'rus', True, window_color_3)
            #         pyautogui.leftClick(color_cord)
            #         pyautogui.sleep(.2)
            pyautogui.leftClick(pyautogui.locateCenterOnScreen(show_spec_button))
            bar.next()
            pyautogui.sleep(.5)
            # pyautogui.leftClick(centerWordSearch("OK", 'eng', False))
            # pyautogui.sleep(.2)
            pyautogui.leftClick(pyautogui.locateCenterOnScreen(one_c_button))
            bar.next()
            pyautogui.sleep(1)
            bar.next()
            pyautogui.leftClick(centerWordSearch("OK", 'eng', False))
            pyautogui.sleep(.2)
            bar.next()
            pyautogui.leftClick(pyautogui.locateCenterOnScreen(exit_button))
            pyautogui.sleep(.2)
            # pyautogui.leftClick(pyautogui.locateCenterOnScreen(exit_button))
            # pyautogui.sleep(.2)
            out_tab["Статус"][index_value] = "Succes"
            out_tab["Ссылка на json"][index_value] = pyperclip.paste()
            bar.next()
            pyautogui.leftClick(pyautogui.locateCenterOnScreen(config_set_button))
            pyautogui.sleep(2)
            bar.finish()
        except Exception:
            e = sys.exc_info()[1]
            out_tab["Статус"][index_value] = e.args[0]
        # create excel writer
        writer = pd.ExcelWriter(r'..\logs.xlsx')
        # write dataframe to excel sheet named 'marks'
        out_tab.to_excel(writer, 'marks')
        # save the excel file
        writer.save()

    # pyautogui.sleep(.2)
    # print(row[1] + detect(row[1]))
    # if len(row[1].split(" ")) == 1:
    #     if detect(row[1]) == 'ru' or detect(row[1]) == 'uk' or detect(row[1]) == 'bk'\
    #             or detect(row[1]) == 'bg' or detect(row[1]) == 'mk':
    #         pyautogui.leftClick(centerWordSearch(row[1], 'rus'))
    #     # elif detect(row[0]) == 'en' or detect(row[0]) == 'cy':
    #     else:
    #         pyautogui.leftClick(centerWordSearch(row[1], 'eng'))
    # elif len(row[1].split(" ")) == 2:
    #     words = row[1].split(" ")
    #     if detect(row[1]) == 'ru' or detect(row[1]) == 'uk' or detect(row[1]) == 'bk'\
    #             or detect(row[1]) == 'bg' or detect(row[1]) == 'mk':
    #         pyautogui.leftClick(neerestWord(words[0], words[1], 'rus'))
    #     # elif detect(row[0]) == 'en' or detect(row[0]) == 'cy':
    #     else:
    #         pyautogui.leftClick(neerestWord(words[0], words[1], 'eng'))
    # pyautogui.sleep(.2)
    # pyautogui.leftClick(pyautogui.locateCenterOnScreen(ok_button))
# print(str(coords[0]) + ', ' + str(coords[1]))
