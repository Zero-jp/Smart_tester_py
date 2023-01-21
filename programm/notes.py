import pyautogui
import time
import progressbar
from colorama import Fore, Back, Style

bar = progressbar.ProgressBar().start()

task_name = pyautogui.locateCenterOnScreen(
    r"C:\Users\drtar\Desktop\SmartCloserForBoasTasks\images\task_name.png", grayscale=False)
move_EndWork = pyautogui.locateCenterOnScreen(
    r"C:\Users\drtar\Desktop\SmartCloserForBoasTasks\images\move_EndWork.png", grayscale=False)
set_EndWork = pyautogui.locateCenterOnScreen(
    r"C:\Users\drtar\Desktop\SmartCloserForBoasTasks\images\set_EndWork.png", grayscale=False)
bad_EndWork_date_1 = pyautogui.locateCenterOnScreen(
    r"C:\Users\drtar\Desktop\SmartCloserForBoasTasks\images\BadEndWorkDate.png", grayscale=False)
bad_EndWork_date_2 = pyautogui.locateCenterOnScreen(
    r"C:\Users\drtar\Desktop\SmartCloserForBoasTasks\images\BadEndWorkDate(2).png", grayscale=False)
tool_bar_open = pyautogui.locateCenterOnScreen(
    r"C:\Users\drtar\Desktop\SmartCloserForBoasTasks\images\tool_bar_open.png", grayscale=False)
wait_please = pyautogui.locateCenterOnScreen(
    r"C:\Users\drtar\Desktop\SmartCloserForBoasTasks\images\wait_please.png", grayscale=False)

# pyautogui.screenshot("img/screen.png")
# start = pyautogui.prompt(title="Начать закрытие?")
# print(pyautogui.locateCenterOnScreen(
#     r"C:\Users\drtar\Desktop\SmartCloserForBoasTasks\images\test_edit.png", grayscale=False))
c = 0
while True:
    if task_name:
        tool_bar_open = pyautogui.locateCenterOnScreen(
            r"C:\Users\drtar\Desktop\SmartCloserForBoasTasks\images\tool_bar_open.png", grayscale=False)
        pyautogui.click(tool_bar_open)
        time.sleep(1)
        # if pyautogui.locateCenterOnScreen(
        #         r"C:\Users\drtar\Desktop\SmartCloserForBoasTasks\images\BadEndWorkDate.png", grayscale=False) \
        #         or pyautogui.locateCenterOnScreen(
        #         r"C:\Users\drtar\Desktop\SmartCloserForBoasTasks\images\BadEndWorkDate(2).png", grayscale=False):
        #     move_EndWork = pyautogui.locateCenterOnScreen(
        #         r"C:\Users\drtar\Desktop\SmartCloserForBoasTasks\images\move_EndWork.png", grayscale=False)
        #     pyautogui.click(move_EndWork)
        #     while pyautogui.locateCenterOnScreen(
        #             r"C:\Users\drtar\Desktop\SmartCloserForBoasTasks\images\BadEndWorkDate(2).png",
        #             grayscale=False) and pyautogui.locateCenterOnScreen(
        #             r"C:\Users\drtar\Desktop\SmartCloserForBoasTasks\images\BadEndWorkDate(2).png", grayscale=False):
        #         time.sleep(2)
        #     time.sleep(3)
        #     pyautogui.click(tool_bar_open)
        #     time.sleep(1)
        #     set_EndWork = pyautogui.locateCenterOnScreen(
        #         r"C:\Users\drtar\Desktop\SmartCloserForBoasTasks\images\set_EndWork.png", grayscale=False)
        #     pyautogui.click(set_EndWork)
        #     while (pyautogui.locateCenterOnScreen(
        #             r"C:\Users\drtar\Desktop\SmartCloserForBoasTasks\images\BadEndWorkDate(2).png",
        #             grayscale=False) is not None) and (pyautogui.locateCenterOnScreen(
        #             r"C:\Users\drtar\Desktop\SmartCloserForBoasTasks\images\BadEndWorkDate(2).png", grayscale=False)
        #                                                is not None):
        #         time.sleep(2)
        #     time.sleep(6)
        # else:
        set_EndWork = pyautogui.locateCenterOnScreen(
            r"C:\Users\drtar\Desktop\SmartCloserForBoasTasks\images\set_EndWork.png", grayscale=False)
        pyautogui.click(set_EndWork)
        time.sleep(6)
        c += 1
        print(Fore.RED + 'Закрыто задач: ' + str(c) + ' || ' + 'Осталось:' + str(264 - c))
    else:
        break

# for t in range(0,101):
#     bar.update(t)
#     time.sleep(0.002)
# bar.finish()
#
# secs = int(input(Fore.GREEN + '              Сколько работаю босс? : '))
# secs = secs * 5
#
# click = 1
# print(Fore.RED + 'Я начну работать через 3 сек!!!')
#
# print(Fore.MAGENTA + str(    1))
# time.sleep(1)
# print(Fore.BLUE + str(    2))
# time.sleep(1)
# print(Fore.BLACK + str(    3))
# time.sleep(1)
#
# for click in range(0, secs):
#     print(Fore.RED + 'До конца осталось >->->->-> ' + str(+ secs))
#     secs -= 1
#     pyautogui.tripleClick()
#     time.sleep(1)
#     time.sleep(0.1)
#     pyautogui.tripleClick()
#     time.sleep(0.1)
#     pyautogui.tripleClick()
#     bar.finish()
#
# else:
#     print('            Я --> з а к о н ч и л --> Б О С С!!!')

from PIL import Image
import pytesseract
import cv2
import os

image = '/tmp/tests.png'

preprocess = "thresh"

# загрузить образ и преобразовать его в оттенки серого
image = cv2.imread(image)
gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

# проверьте, следует ли применять пороговое значение для предварительной обработки изображения

if preprocess == "thresh":
    gray = cv2.threshold(gray, 0, 255,
                         cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]

# если нужно медианное размытие, чтобы удалить шум
elif preprocess == "blur":
    gray = cv2.medianBlur(gray, 3)

# сохраним временную картинку в оттенках серого, чтобы можно было применить к ней OCR

filename = "{}.png".format(os.getpid())
cv2.imwrite(filename, gray)
# загрузка изображения в виде объекта image Pillow, применение OCR, а затем удаление временного файла
text = pytesseract.image_to_string(Image.open(filename))
os.remove(filename)
print(text)
# показать выходные изображения
cv2.imshow("Image", image)
cv2.imshow("Output", gray)
# input(‘pause…’)
