import pyautogui
import win32api, win32con
import time
import pytesseract # оптический анализатор(нейронка)
import pandas as pd # считывание из таблицы входных данных
import math
import statistics
from PIL import Image, ImageFilter
from langdetect import detect

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


# def leftClick(coord):
#     # клик
#     win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, coord[0], coord[1])
#     # time.sleep(.1)
#     win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, coord[0], coord[1])
#     print("Click.")
def lookOnScreen(lang: str):
    # поиск слова на экране
    temp_screen = pyautogui.screenshot('temp.bmp')
    img_seryy = temp_screen.convert("L").save(r'C:\Users\drtar\Desktop\SmartTesterForBoas\programm\temp_kraya.bmp')
    # temp_kraya = img_seryy.filter(ImageFilter.FIND_EDGES)#.save(r'C:\Users\drtar\Desktop\SmartTesterForBoas\programm\temp_kraya.bmp')
    #temp_kraya = img_seryy.filter(ImageFilter.SHARPEN)
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
    if lang != 'eng':
        data = pytesseract.image_to_data(r'C:\Users\drtar\Desktop\SmartTesterForBoas\programm\temp_kraya.bmp', lang=lang, output_type=pytesseract.Output.DICT)
    else:
        data = pytesseract.image_to_data(r'C:\Users\drtar\Desktop\SmartTesterForBoas\programm\temp_kraya.bmp',
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
    data = lookOnScreen(lang)
    # получение координат
    print(word)
    print(data["text"])
    for d in data["text"]:
        if d.find(word) != -1:
            print(word)
            print(d)
            word_id = data["text"].index(d)
            x = data["left"][word_id] + data["width"][word_id] / 2
            y = data["top"][word_id] + data["height"][word_id] / 2
    return [int(x), int(y)]


def dist(p1, p2):
    return math.hypot(p2[0] - p1[0], p2[1] - p1[1])


def neerestFigure(link: str, coords: list):
    # получение координат
    minDist = 1000
    imageCoords = [0, 0]
    for pos in pyautogui.locateAllOnScreen(link):
        # поиск ближайшей точки
        d = dist(pos, coords)
        if d < minDist:
            imageCoords = [pos[0], pos[1]]
            minDist = d
    return imageCoords

def neerestWord(word_main: str, word_sub: str):
    data = lookOnScreen('rus')
    main_coords = []
    sub_coords = []
    meanCoords = [0, 0]
    minDist = 1000
    for i in range(len(data["text"])):
        # print(d)

        if data["text"][i] == word_main and data["text"][i+1] == word_sub:
            print("Main!")
            main_coords = [data["left"][i] + data["width"][i] / 2, data["top"][i] + data["height"][i] / 2]
            sub_coords = [data["left"][i+1] + data["width"][i+1] / 2, data["top"][i+1] + data["height"][i+1] / 2]
    meanCoords = [(main_coords[0] + sub_coords[0]) / 2, (main_coords[1] + sub_coords[1]) / 2]
    print(meanCoords)
    return meanCoords

# Забираем данные для прохода по полям
data = pd.DataFrame(pd.read_excel(
    io=r'C:\Users\drtar\Desktop\SmartTesterForBoas\Копия (Агрегатор Т1) ТК для тестирования интеграции между ELMA и IMOS (3).xlsx',
    engine='openpyxl'))
frame = r'C:\Users\drtar\Desktop\SmartTesterForBoas\images\frame.bmp'
ok_button = r'C:\Users\drtar\Desktop\SmartTesterForBoas\images\ok_button.bmp'

# print(data.columns[0])
colum_name = []
colum_name.append(data.columns[0].split(' '))
colum_name.append(data.columns[1].split(' '))
# print(colum_name[0][0])
# print(colum_name[0][1])
# pyautogui.leftClick(neerestWord(colum_name[0][0], colum_name[0][1]))
for index, row in data.iterrows():
    # print(row[0] + '|' + row[1])
    first_name = neerestWord(colum_name[0][0], colum_name[0][1])
    pyautogui.leftClick(neerestFigure(frame, first_name))
    pyautogui.sleep(.1)
    listWordSearch()
    # print(row[0])
    # print(detect(row[0]))
    if detect(row[0]) == 'ru':
        pyautogui.leftClick(centerWordSearch(row[0], 'rus'))
    elif detect(row[0]) == 'en' or detect(row[0]) == 'cy':
        pyautogui.leftClick(centerWordSearch(row[0], 'eng'))
    pyautogui.sleep(.1)
    second_name = neerestWord(colum_name[1][0], colum_name[1][1])
    pyautogui.leftClick(neerestFigure(frame, second_name))
    pyautogui.sleep(.1)
    pyautogui.leftClick(centerWordSearch(row[1]))
    pyautogui.sleep(.1)
    # pyautogui.leftClick(pyautogui.locateCenterOnScreen(ok_button))
# print(str(coords[0]) + ', ' + str(coords[1]))
