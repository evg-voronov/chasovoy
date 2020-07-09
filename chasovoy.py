import cv2
import openpyxl
import winsound
from datetime import datetime, timedelta
import time
import os
import random


def audio_message(path, time_start, hour):
    if time_start is None:
        hour = 1
    if time_here // 3600 == hour:
        wav = random.choice(os.listdir(path + r'\audio'))
        winsound.PlaySound(path + r'\audio\\' + wav, winsound.SND_FILENAME)
        hour = time_here // 3600 + 1
    return hour


def show_face_frame():
    if len(faces) > 0:
        for (x, y, w, h) in faces:
            cv2.rectangle(img, (x, y), (x+w, y+h), (0, 255, 0), 2)


def write_excel(path_excel, start):
    wb = openpyxl.load_workbook(path_excel + r'\dnevnik.xlsx')
    today = datetime.strftime(start, '%d.%m.%Y')
    if today not in wb.sheetnames:  # если листа с сегодняшней датой еще не существует, то создаем его
        wb.create_sheet(title=today)
        sheet = wb[today]
        sheet.column_dimensions['A'].width = 20
        sheet.column_dimensions['B'].width = 20
        sheet.column_dimensions['C'].width = 20
        sheet.column_dimensions['D'].width = 31
        sheet['A1'] = 'Начал работать за ПК'
        sheet['B1'] = 'Время работы за ПК'
        sheet['C1'] = 'Закончил работать'
        sheet['D1'] = 'Сели за компьютер:'
        sheet['D2'] = 'Всего провели за компьютером:'
        sheet = wb[today]
        row = 2
        all_time = 0
    else:  # если лист уже создан извлекаем информацию
        sheet = wb[today]
        row = sheet['E1'].value  # номер строки в excel документе
        all_time = sheet['E2'].value  # общее время проведенное за ПК за один день
        all_time = all_time.split(':')
        all_time = int(all_time[0]) * 3600 + int(all_time[1]) * 60 + int(all_time[2])  # время переводим в секунды
        row = row + 2  # плюсуем шапку таблицы и новую строку

    sheet['A' + str(row)] = datetime.strftime(start, '%H:%M:%S')  # записываем время когда сели за пк
    sheet['C' + str(row)] = datetime.strftime(datetime.now(), '%H:%M:%S')  # когда вышли из за пк
    print('закончили работать в ' + datetime.strftime(datetime.now(), '%H:%M:%S'))
    time_count = datetime.now() - start  # время которое мы провели за компьютером
    seconds = time_count.seconds
    print('вы провели за компьютером ' + time.strftime('%H:%M:%S', time.gmtime(seconds)))
    sheet['B' + str(row)] = time.strftime('%H:%M:%S', time.gmtime(seconds))
    sheet['E1'] = row - 1  # вычитаем шапку таблицу
    all_time = all_time + seconds
    all_time = time.strftime('%H:%M:%S', time.gmtime(all_time))
    sheet['E2'] = all_time
    # заполняем первый лист
    sheet_0 = wb.worksheets[0]
    sheet_0['A' + str(len(wb.sheetnames))] = today
    sheet_0['B' + str(len(wb.sheetnames))] = all_time
    while True:  # закрываем файл
        try:
            wb.save(path + r'\dnevnik.xlsx')
        except PermissionError:
            input('Сначала закройте excel файл и подтвердите нажав enter')
        else:
            break


cap = cv2.VideoCapture(0)
path = r'C:\Users\Voronov\YandexDisk-euge.voronov@yandex.ru\Manual\Project\chasovoy'
faceCascade = cv2.CascadeClassifier(path + r'\haarcascade_frontalface_default.xml')

time_here, time_not_here = 0, 0  # подсчет времени когда человек за компьютером и когда его там нет.
here, not_here = 0, 0
min_time = 60
time_start = None
status = True
hour = 1

try:
    while True:
        ret, img = cap.read()
        if ret == 0:
            break
        img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        faces = faceCascade.detectMultiScale(img_gray, scaleFactor=1.3, minNeighbors=3, minSize=(30, 30))

        show_face_frame()
        hour = audio_message(path, time_start, hour)

        if len(faces) > 0 and status:
            not_here = 0
            status = False
            if time_start is None:
                here = time.time()
                time_start = datetime.now()
                print('вы сели за компьютер в ' + datetime.strftime(time_start, '%H:%M:%S'))
        elif len(faces) == 0 and not status:
            not_here = time.time()
            status = True

        if not_here != 0:
            time_not_here = time.time() - not_here
        if here != 0 and not status:
            time_here = time.time() - here
        # print(str(time_here) + 'here')
        # print(str(time_not_here) + 'not_here')
        # time.sleep(0.2)

        if time_here > min_time and time_not_here > min_time:  # делаем запись если провели за ПК больше минуты
            write_excel(path, time_start)
            time_here, time_not_here = 0, 0
            not_here = 0
            time_start = None
        elif time_not_here > min_time > time_here:  # если за ПК провели времени меньше минуты, то ничего не записываем
            print('запись не произведена')
            time_here, time_not_here = 0, 0
            not_here = 0
            time_start = None

        # cv2.imshow('frame', img)
        cv2.waitKey(1)
except KeyboardInterrupt:
    if time_here > min_time:  # делаем запись если провели за ПК больше минуты
        write_excel(path, time_start)
    cap.release()
    cv2.destroyAllWindows()
