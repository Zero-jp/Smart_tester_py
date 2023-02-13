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

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
Column_Coords = [0, 0]

# def leftClick(coord):
#     # клик
#     win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, coord[0], coord[1])
#     # time.sleep(.1)
#     win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, coord[0], coord[1])
#     print("Click.")
def lookOnScreen(lang: str, is_obrez: bool):
    # поиск слова на экране
    # 1018 294  1205 536
    temp_screen = pyautogui.screenshot('temp.bmp')
    img_seryy = temp_screen.convert("L")# .save(r'C:\Users\drtar\Desktop\SmartTesterForBoas\programm\temp_kraya.bmp')
    # temp_kraya = img_seryy.filter(ImageFilter.FIND_EDGES)#.save(r'C:\Users\drtar\Desktop\SmartTesterForBoas\programm\temp_kraya.bmp')
    img_seryy.filter(ImageFilter.DETAIL).save(r'C:\Users\drtar\Desktop\SmartTesterForBoas\programm\temp_kraya.bmp')
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
    print(lang)
    link = r'C:\Users\drtar\Desktop\SmartTesterForBoas\programm\temp_kraya.bmp'
    link_obrez = r'C:\Users\drtar\Desktop\SmartTesterForBoas\programm\temp_kraya_obrez.bmp'
    if is_obrez:
        print(Column_Coords)
        temp_kraya = Image.open(link)
        im_crop = temp_kraya.crop((Column_Coords[0], 0, Column_Coords[1], GetSystemMetrics(1)))
        im_crop.save(link_obrez)
    l = ""
    if is_obrez:
        l = link_obrez
    else:
        l = link
    if lang != 'eng':
        data = pytesseract.image_to_data(l, lang=lang, config='--oem 3 --psm 12', output_type=pytesseract.Output.DICT)
    else:
        data = pytesseract.image_to_data(l, lang=lang, config='--oem 3 --psm 12',
                                         output_type=pytesseract.Output.DICT)
    print(data)

    return data


def listWordSearch():
    temp_screen = pyautogui.screenshot('temp.bmp')
    img_seryy = temp_screen.convert("L")
    list_star = pyautogui.locateCenterOnScreen(r'C:\Users\drtar\Desktop\SmartTesterForBoas\images\list_start.bmp')
    print(list_star)
    pyautogui.leftClick(list_star)


def centerWordSearch(word: str, lang: str):
    # поиск слова на экране
    # получить все данные из изображения
    data = lookOnScreen(lang, True)
    # получение координат
    #print("искомое слово:" + word)
    #print(data["text"])
    for d in data["text"]:
        if d.find(word) != -1:
            #print(word)
            print(d)
            word_id = data["text"].index(d)
            x = data["left"][word_id] + data["width"][word_id] / 2
            y = data["top"][word_id] + data["height"][word_id] / 2
    return [int(x)+Column_Coords[0], int(y)]


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
    print(imageCoords)
    return imageCoords

def neerestWord(word_main: str, word_sub: str, lang: str):
    data = lookOnScreen(lang, False)
    main_coords = []
    sub_coords = []
    meanCoords = [0, 0]
    minDist = 1000
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

# Забираем данные для прохода по полям
data = pd.DataFrame(pd.read_excel(
    io=r'C:\Users\drtar\Desktop\SmartTesterForBoas\Копия (Агрегатор Т1) ТК для тестирования интеграции между ELMA и IMOS (3).xlsx',
    engine='openpyxl'))
out_tab = data.copy()
frame = r'C:\Users\drtar\Desktop\SmartTesterForBoas\images\frame.bmp'
full_frame = r'C:\Users\drtar\Desktop\SmartTesterForBoas\images\full_frame.bmp'
ok_button = r'C:\Users\drtar\Desktop\SmartTesterForBoas\images\ok_button.bmp'
error_wind = r'C:\Users\drtar\Desktop\SmartTesterForBoas\images\error_wind.bmp'
active_ok_button = r'C:\Users\drtar\Desktop\SmartTesterForBoas\images\active_ok_button.bmp'
text_field = r'C:\Users\drtar\Desktop\SmartTesterForBoas\images\text_field.bmp'
use_config_wind = r'C:\Users\drtar\Desktop\SmartTesterForBoas\images\use_config_wind.bmp'
active_yes_button = r'C:\Users\drtar\Desktop\SmartTesterForBoas\images\active_yes_button.bmp'
show_spec_button = r'C:\Users\drtar\Desktop\SmartTesterForBoas\images\show_spec_button.bmp'
one_c_button = r'C:\Users\drtar\Desktop\SmartTesterForBoas\images\one_c_button.bmp'
exit_button = r'C:\Users\drtar\Desktop\SmartTesterForBoas\images\exit_button.bmp'
config_set_button = r'C:\Users\drtar\Desktop\SmartTesterForBoas\images\config_set_button.bmp'

# print(data.columns[0])
colum_name = []
columns_count = data.shape[0]
rows_count = data.shape[1]
for index in range(columns_count):
    colum_name.append(data.columns[index].split(' '))
# print(colum_name[0][0])
# print(colum_name[0][1])
#pyautogui.leftClick(neerestWord(colum_name[0][0], colum_name[0][1]))

for index, row in data.iterrows():
    try:
        # Нахождение названия свойства и элемента со списком
        #print(row[0] + '|' + row[1])
        first_name = neerestWord(colum_name[0][0], colum_name[0][1], 'rus')
        pyautogui.leftClick(neerestFigure(frame, first_name))
        pyautogui.sleep(.2)
        # listWordSearch()
        # print(row[0])
        # print(detect(row[0]))
        if detect(row[0]) == 'ru':
            pyautogui.leftClick(centerWordSearch(row[0], 'rus'))
        # elif detect(row[0]) == 'en' or detect(row[0]) == 'cy':
        else:
            pyautogui.leftClick(centerWordSearch(row[0], 'eng'))
        pyautogui.sleep(.2)
        print(colum_name[1][0] + '|' + colum_name[1][1])
        second_name = neerestWord(colum_name[1][0], colum_name[1][1], 'rus')
        pyautogui.leftClick(neerestFigure(frame, second_name))
        # Нахождение нужного значения
        for index_value in range(columns_count-1):
            index_value += 1
            print(row[index_value] + detect(row[index_value]))
            if len(row[index_value].split(" ")) == 1:
                if detect(row[index_value]) == 'ru' or detect(row[index_value]) == 'uk' or detect(row[index_value]) == 'bk' \
                        or detect(row[index_value]) == 'bg' or detect(row[index_value]) == 'mk':
                    pyautogui.leftClick(centerWordSearch(row[index_value], 'rus'))
                # elif detect(row[0]) == 'en' or detect(row[0]) == 'cy':
                else:
                    pyautogui.leftClick(centerWordSearch(row[index_value], 'eng'))
            elif len(row[index_value].split(" ")) == 2:
                words = row[index_value].split(" ")
                if detect(row[index_value]) == 'ru' or detect(row[index_value]) == 'uk' or detect(row[index_value]) == 'bk' \
                        or detect(row[index_value]) == 'bg' or detect(row[index_value]) == 'mk':
                    pyautogui.leftClick(neerestWord(words[0], words[1], 'rus'))
                # elif detect(row[0]) == 'en' or detect(row[0]) == 'cy':
                else:
                    pyautogui.leftClick(neerestWord(words[0], words[1], 'eng'))
            pyautogui.sleep(.2)
            pyautogui.leftClick(pyautogui.locateCenterOnScreen(ok_button))
            pyautogui.sleep(.2)
            if pyautogui.locateCenterOnScreen(error_wind) != None:
                pyautogui.leftClick(pyautogui.locateCenterOnScreen(active_ok_button))
                pyautogui.sleep(.2)
            pyautogui.leftClick(neerestFigure(text_field, "Номер"))
            pyautogui.sleep(.2)
            pyautogui.write(datetime.datetime.strptime(datetime.datetime.now(), '%Y-%m-%d|%H:%M:%S.%f'))
            pyautogui.leftClick(pyautogui.locateCenterOnScreen(ok_button))
            pyautogui.sleep(.2)
            if pyautogui.locateCenterOnScreen(use_config_wind) != None:
                pyautogui.leftClick(pyautogui.locateCenterOnScreen(active_yes_button))
                pyautogui.sleep(.2)
            pyautogui.leftClick(pyautogui.locateCenterOnScreen(show_spec_button))
            pyautogui.sleep(.5)
            pyautogui.leftClick(pyautogui.locateCenterOnScreen(one_c_button))
            pyautogui.sleep(1)
            pyautogui.leftClick(pyautogui.locateCenterOnScreen(exit_button))
            pyautogui.sleep(.2)
            pyautogui.leftClick(pyautogui.locateCenterOnScreen(exit_button))
            pyautogui.sleep(.2)
        out_tab["Статус"][index_value] = "Succes"
        out_tab["Ссылка на json"][index_value] = pyperclip.paste()
        pyautogui.leftClick(pyautogui.locateCenterOnScreen(config_set_button))
        pyautogui.sleep(.2)
    except:
        out_tab["Статус"][index_value] = "False"
    # create excel writer
    writer = pd.ExcelWriter(r'C:\Users\drtar\Desktop\SmartTesterForBoas\logs.xlsx')
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
