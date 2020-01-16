"""
Создает PNG обложки для печати на CD-дисках
в соответствии с переданным DOCX файлом реферата
"""

import io
import os
import sys
import logging  # Для логирования
import argparse  # Для работы с аргументами командной строки
import graphics  # Отрисовка графики
import docx  # Работа с .docx файлами
from PIL import Image


def is_windows_64bit():
    """ 
        Определяет архитектуру ОС.
        Возвращает True, если архитектура x64.
    """
    if "PROCESSOR_ARCHITEW6432" in os.environ:
        return True
    return os.environ["PROCESSOR_ARCHITECTURE"].endswith("64")


def main(file, title_len=60, title_size=8):
    """
    Заполняет словарь с данными из реферата 
    и формирует обложки для подлинника и дубликата CD
    file - имя файла реферата
    title_len - кол-во символов на строку для названия CD
    title_size - размер шрифта названия CD
    """

    def draw_text(win, point, text_string, text_size):
        """ 
        Отрисовывает текст
        win - graphics.GraphWin - холст, на котором будет отрисован текст 
        point - graphics.Point - координаты центра прямоугольника в котором
                             будет отрисован текст
        text_string - str           - текст для отрисовки ()
        text_size - int           - размер шрифта текста
        """
        
        text = graphics.Text(point, text_string)
        text.setSize(text_size)
        text.draw(win)

    try:
        document = docx.Document(file)  # Открытие docx файла
    except Exception as err:
        logging.error(str(err))
        sys.exit()

    labels = dict()  # Словарь с данными из считанного docx файла

    table = document.tables[0]  # Вынимаем таблицу с данными

    # Перебираем строки таблицы table и заполняем текстом из полей таблицы словарь labels
    for row in table.rows:
        # Убираем символы переноса строк из полей "Ключ"
        key = row.cells[1].text.replace("\n", " ")
        while key in labels.keys():
            key = key + "_"
        labels[key] = row.cells[2].text

    volumes = 1  # Базовое кол-во томов

    # Список словарей для каждого тома
    disks = []

    # Находим необходимые строки в таблице реферата
    for key in labels:
        if "регистрационный" in key.lower():
            regNum = key
            continue
        elif "название" in key.lower():
            name = key
            continue
        elif "децимальный" in key.lower():
            decimalNum = key
            continue
        elif "рассылка" in key.lower():
            cdType = key
            continue
        elif "контрольная" in key.lower():
            ksum = key

    # Проверяем, что все значения для полей диска присутствуют в реферате
    try:
        if regNum and name and decimalNum and cdType and ksum:
            pass
    except Exception as err:
        logging.error(str(err))
        sys.exit()

    while volumes:

        disks.append(dict())  # Добавляем новый словарь для тома

        # Заполнение словаря тома
        disks[volumes - 1][regNum] = labels[regNum]

        disks[volumes - 1][name] = labels[name] + (
            "." if labels[name][-1] != "." else "")
            
        if "Том " + str(volumes) + ":" in labels:
            disks[volumes - 1][name] = (
                disks[volumes - 1][name]
                + " "
                + labels["Том " + str(volumes) + ":"]
                + ("." if labels["Том " + str(volumes) + ":"][-1] != "." else ""))

        disks[volumes - 1][decimalNum] = labels[decimalNum]
        if "Том " + str(volumes) + ":_" in labels:
            disks[volumes - 1][decimalNum] = (
                disks[volumes - 1][decimalNum] + labels["Том " + str(volumes) + ":_"])

        # Находим тип носителя (CD, DVD и т.п.)
        words = labels[cdType].split()
        for word in words:
            if "D" in word:
                s = word
                break

        disks[volumes - 1][cdType] = s

        disks[volumes - 1][ksum] = labels[ksum]
        if "Том " + str(volumes) + ":___" in labels:
            disks[volumes - 1][ksum] = (
                disks[volumes - 1][ksum] + labels["Том " + str(volumes) + ":___"])

        # Если находим метку "Том Х:" - увеличиваем кол-во томов к печати
        if "Том " + str(volumes + 1) + ":" in labels:
            volumes += 1
        else:
            volumes = 0

    # Формируем обложку для каждого тома в 2х экземплярах - подлинник и дубликат
    for disk in disks:

        # Инициализируем холст
        win = graphics.GraphWin("Окно для графики", 448, 448)

        # Большая окружность
        obj = graphics.Circle(graphics.Point(224, 224), 224)
        obj.setOutline("white")
        obj.setFill("#87CEFA")
        obj.draw(win)

        # Малая окружность
        obj = graphics.Circle(graphics.Point(224, 224), 80)
        obj.setOutline("white")
        obj.setFill("white")
        obj.draw(win)

        # Прямоугольник для названия организации
        obj = graphics.Rectangle(graphics.Point(104, 36), graphics.Point(344, 55))
        obj.draw(win)

        # Текст названия организации
        draw_text(win, graphics.Point(224, 45), "ФГУП «ГосНИИАС»", 10)

        # Прямоугольник для названия CD
        obj = graphics.Rectangle(graphics.Point(80, 55), graphics.Point(368, 135))
        obj.draw(win)

        # Текст названия CD
        # Заменяем '.' на '.'+'\n'
        s = disk[name].split(".")
        s = ".\n".join(s)

        # Разбиваем строку на слова по пробельному символу (и символу новой строки)
        string = s.split()

        # Ограничиваем длину строки title_len(60 по дефолту) символами, 
        # после чего начинаем печать на новой строке
        val = ""  # Строка с ограничением в title_len символов
        n = 1  # Счетчик перехода на новую строку
        for word in string:
            if len(val) + len(word) > title_len * n:
                val += "\n"
                n += 1
            val += word
            val += " "

        # Текст "Название программы" disk["Название программы/документа/документации:"])
        draw_text(win, graphics.Point(224, 85), val, title_size)

        # Текст с децимальным номером из реферата
        draw_text(win, graphics.Point(224, 125), disk[decimalNum], 10)

        # Прямоуголник "Вид носителя"
        obj = graphics.Rectangle(graphics.Point(20, 165), graphics.Point(140, 185))
        obj.draw(win)
        obj = graphics.Rectangle(graphics.Point(100, 165), graphics.Point(140, 185))
        obj.draw(win)

        # Текст "Вид носителя"
        draw_text(win, graphics.Point(60, 175), "Вид носителя", 10)

        # Пишем вид носителя
        draw_text(win, graphics.Point(120, 175), disk[cdType], 10)

        # Прямоуголник "Подразделение"
        obj = graphics.Rectangle(graphics.Point(308, 165), graphics.Point(428, 185))
        obj.draw(win)
        obj = graphics.Rectangle(graphics.Point(388, 165), graphics.Point(428, 185))
        obj.draw(win)

        # Текст "Подразделение"
        draw_text(win, graphics.Point(348, 175), "Подразделение", 10)

        # Текст "0500"
        draw_text(win, graphics.Point(408, 175), "0500", 10)

        # Прямоугольник "Контрольная характеристика"
        obj = graphics.Rectangle(graphics.Point(15, 215), graphics.Point(135, 295))
        obj.draw(win)
        obj = graphics.Rectangle(graphics.Point(15, 215), graphics.Point(135, 240))
        obj.draw(win)

        # Текст "Контрольная характеристика"
        draw_text(win, graphics.Point(75, 230), "Контрольная характеристика", 8)

        # Текст c контрольной суммой из реферата
        # Разбиваем КСумм на две строки
        s = disk[ksum]
        l = len(s)
        s = "\n".join([s[0 : l // 2], s[l // 2 : :]])

        # Пишем Ксумм
        draw_text(win, graphics.Point(75, 265), s, 10)

        # Прямоугольник с подписями ВП, ОТК
        obj = graphics.Rectangle(graphics.Point(314, 215), graphics.Point(432, 295))
        obj.draw(win)
        obj = graphics.Rectangle(graphics.Point(314, 215), graphics.Point(432, 235))
        obj.draw(win)
        obj = graphics.Rectangle(graphics.Point(314, 215), graphics.Point(432, 255))
        obj.draw(win)
        obj = graphics.Rectangle(graphics.Point(314, 215), graphics.Point(432, 275))
        obj.draw(win)
        obj = graphics.Rectangle(graphics.Point(314, 215), graphics.Point(360, 295))
        obj.draw(win)

        # Текст "ОТК", "ВП МО", "Дата"
        draw_text(win, graphics.Point(337, 225), "ОТК", 10)
        draw_text(win, graphics.Point(337, 245), "Дата", 10)
        draw_text(win, graphics.Point(337, 265), "ВП МО", 10)
        draw_text(win, graphics.Point(337, 285), "Дата", 10)

        # Прямоугольник для номера тома, рег. номера и вида ЭД
        obj = graphics.Rectangle(graphics.Point(78, 350), graphics.Point(370, 390))
        obj.draw(win)
        obj = graphics.Rectangle(graphics.Point(78, 350), graphics.Point(370, 370))
        obj.draw(win)
        obj = graphics.Rectangle(graphics.Point(150, 350), graphics.Point(300, 390))
        obj.draw(win)

        # Текст "Номер тома/Количество томов"
        draw_text(win, graphics.Point(114, 355), "Номер тома/", 8)
        draw_text(win, graphics.Point(114, 365), "Количество томов", 8)

        # Пишем номер тома
        t = f"{disks.index(disk)+1}/{len(disks)}"
        draw_text(win, graphics.Point(114, 380), t, 8)

        # Текст "Регистрационный номер"
        draw_text(win, graphics.Point(225, 360), "Регистрационный номер", 8)

        # Вставляем регистрационный номер из реферата
        draw_text(win, graphics.Point(225, 380), disk[regNum], 8)

        # Текст "Вид ЭД"
        draw_text(win, graphics.Point(335, 360), "Вид ЭД", 8)

        # Пишем, что формируем подлинник
        draw_text(win, graphics.Point(335, 380), "П", 8)

        # Сохранение подлинника
        # Преобразуем изображение в формат postscript
        ps = win.postscript(colormode="color", pageheight=448, pagewidth=448)
        # win.postscript(colormode='color', file="image.eps")
        # Кодируем преобразованное изображение в байты
        img = Image.open(io.BytesIO(ps.encode("utf-8")))
        # Чтобы не потерять качество при преобразовании в .png увеличиваем
        # postscript изображение в 8 раз
        img.load(scale=8)
        # Возвращаем исходный масштаб
        img = img.resize((1394, 1394), Image.BICUBIC)
        # Сохраняем изображение
        img.save("label" + str(disks.index(disk)) + "_p.png", "png", dpi=(300, 300))

        # Ждем клика мышью на окне
        # win.getMouse()

        # Закрашиваем прямоугольник "Контрольная характеристика" для дубликата
        obj = graphics.Rectangle(graphics.Point(15, 215), graphics.Point(135, 295))
        obj.draw(win)
        obj.setFill("#87CEFA")
        obj = graphics.Rectangle(graphics.Point(15, 215), graphics.Point(135, 240))
        obj.draw(win)

        # Текст "Контрольная характеристика"
        draw_text(win, graphics.Point(75, 230), "Контрольная характеристика", 8)

        # Закрашиваем прямоугольник "Вид ЭД"
        obj = graphics.Rectangle(graphics.Point(300, 370), graphics.Point(370, 390))
        obj.setFill("#87CEFA")
        obj.draw(win)

        # Пишем, что формируем дубликат
        draw_text(win, graphics.Point(335, 380), "Д", 8)

        # Ждем клика мышью на окне
        # win.getMouse()

        # Сохранение дубликата
        # Преобразуем изображение в формат postscript
        ps = win.postscript(colormode="color", pageheight=448, pagewidth=448)
        # win.postscript(colormode='color', file="image.eps")
        # Кодируем преобразованное изображение в байты
        img = Image.open(io.BytesIO(ps.encode("utf-8")))
        # Чтобы не потерять качество при преобразовании в .png увеличиваем
        # postscript изображение в 8 раз
        img.load(scale=8)
        # Возвращаем исходный масштаб
        img = img.resize((1394, 1394), Image.BICUBIC)
        # Сохраняем изображение
        img.save("label" + str(disks.index(disk)) + "_d.png", "png", dpi=(300, 300))

        #  Закрываем окно
        win.close()


# Включаем протоколирование ошибок и сообщений
logging.basicConfig(
    filename="log.txt",
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

parser = argparse.ArgumentParser(description="Referat file to CD cover")
parser.add_argument("file", type=str, help="Path to referat file")
parser.add_argument(
    "-l", dest='title_len', type=int, default=60, help="Maximum characters per title string, default=60"
)
parser.add_argument("-f", dest='title_size', type=int, default=8, help="Title font size, default=8")
args = parser.parse_args()
# print(args.__dict__)

# Добавляем путь в SYSTEM PATH до Ghostscript под нужную архитектуру
path = os.path.dirname(
    os.path.abspath(args.file)
)  # Абсолютный путь до папки со скриптом label.py

# Определяем системную архитектуру и дополняем путь до папки с Ghostscript в зависимости от архитектуры
if is_windows_64bit():
    path = path + r"\x64"
else:
    path = path + r"\x32"

# Модифицируем PATH
app_path = os.path.join(path)
os.environ["PATH"] += os.pathsep + app_path

logging.debug(" ".join(str(args.__dict__)))

# Запуск основного скрипта
main(*args.__dict__.values())
