import graphics  # Отрисовка графики
import sys       # Работа с аргументами, переданными скрипту
import docx      # Работа с .docx файлами 
import io
import os
import time
from PIL import Image


def is_windows_64bit():
    if 'PROCESSOR_ARCHITEW6432' in os.environ:
        return True
    return os.environ['PROCESSOR_ARCHITECTURE'].endswith('64')


def main(file, title_len=60, title_size=8):

    try:
        document = docx.Document(file)  # Открытие docx файла
    except Exception as error:
        print(f"Can't open {file}")
        sys.exit()
  
    labels = dict()   # Словарь с данными из считанного docx файла
    
    table = document.tables[0]  # Вынимаем таблицу с данными
    
    # for row in table.rows:
        # print(row.cells[1].text + '\t' + row.cells[2].text)
    
    
    # Перебираем строки таблицы table и заполняем текстом из полей таблицы словарь labels
    for row in table.rows:
        # Убираем символы переноса строк из полей "Ключ"
        key = row.cells[1].text.replace('\n', ' ')
        while key in labels.keys():
            key = key + '_'
        labels[key] = row.cells[2].text
         

    volumes = 1 # Базовое кол-во томов

    # Список словарей для каждого тома
    disks = []

    while volumes:

        disks.append(dict())    # Добавляем новый словарь для тома

        # Заполнение словаря тома 
        disks[volumes-1]["Регистрационный номер:"] = labels["Регистрационный номер:"]
        
        disks[volumes-1]["Название программы/документа/документации:"] = labels["Название программы/документа/документации:"] + ("." if labels["Название программы/документа/документации:"][-1] != "." else "")        
        if "Том " + str(volumes) + ":" in labels.keys():
            disks[volumes-1]["Название программы/документа/документации:"] = disks[volumes-1]["Название программы/документа/документации:"]+ " " + labels["Том " + str(volumes) + ":"] + ("." if labels["Том " + str(volumes) + ":"][-1] != "." else "")

        disks[volumes-1]["Децимальный номер (УДК): № версии, редакции:"] = labels["Децимальный номер (УДК): № версии, редакции:"]        
        if "Том " + str(volumes) + ":_" in labels.keys():
            disks[volumes-1]["Децимальный номер (УДК): № версии, редакции:"] = disks[volumes-1]["Децимальный номер (УДК): № версии, редакции:"] + labels["Том " + str(volumes) + ":_"]

        # Находим тип носителя (CD, DVD и т.п.)
        words = labels["Рассылка (общее кол-во дисков) (шт.):"].split()
        for word in words:
            if "D" in word:
                s = word
                break
        
        disks[volumes-1]["Рассылка (общее кол-во дисков) (шт.):"] = s

        disks[volumes-1]["Контрольная характеристика (сумма)"] = labels["Контрольная характеристика (сумма)"]        
        if "Том " + str(volumes) + ":___" in labels.keys():
            disks[volumes-1]["Контрольная характеристика (сумма)"] = disks[volumes-1]["Контрольная характеристика (сумма)"] + labels["Том " + str(volumes) + ":___"]

        if "Том " + str(volumes+1) + ":" in labels.keys():      # Если находим метку "Том Х:" - увеличиваем кол-во томов к печати
            volumes += 1
        else:
            volumes = 0


    # Формируем обложку для каждого тома в 2х экземплярах - подлинних и дубликат
    for disk in disks:
    
        win = graphics.GraphWin("Окно для графики", 448, 448)
    
        # Большая окружность
        obj = graphics.Circle(graphics.Point(224, 224), 224)
        obj.setOutline("white")
        obj.setFill("#87CEFA")
        obj.draw(win)
    
        # Малая окружность
        obj = graphics.Circle(graphics.Point(224,224), 80)
        obj.setOutline("white")
        obj.setFill("white")
        obj.draw(win)
        
    
        # Прямоугольник для названия организации
        obj = graphics.Rectangle(graphics.Point(104, 36), graphics.Point(344, 55))
        obj.draw(win)
     
        # Текст названия организации   
        text = graphics.Text(graphics.Point(224,45),"ФГУП «ГосНИИАС»")
        text.setSize(10)
        text.draw(win)
    
    
        # Прямоугольник для названия CD
        obj = graphics.Rectangle(graphics.Point(80, 55), graphics.Point(368, 135))
        obj.draw(win)
    
        # Текст названия CD
        # Заменяем '.' на '.'+'\n'
        s = disk["Название программы/документа/документации:"].split('.')
        s = '.\n'.join(s)

        # Разбиваем строку на слова по пробельному символу (и символу новой строки)
        string = s.split()

        # Ограничиваем длину строки title_len(60 по дефолту) символами, после чего начинаем печать на новой строке
        val =""     # Строка с ограничением в title_len символов
        n = 1       # Счетчик перехода на новую строку
        for word in string:
            if len(val) + len(word) > int(title_len)*n:
                val += '\n'
                n += 1
            val += word
            val += " "

        # text = graphics.Text(graphics.Point(224,85), disk["Название программы/документа/документации:"])
        text = graphics.Text(graphics.Point(224,85), val)
        text.setSize(int(title_size))
        text.draw(win)
    
        # Текст с децимальным номером из реферата
        text = graphics.Text(graphics.Point(224,125), disk["Децимальный номер (УДК): № версии, редакции:"])
        text.setSize(10)
        text.draw(win)
    
        # Прямоуголник "Вид носителя"
        obj = graphics.Rectangle(graphics.Point(20, 165), graphics.Point(140, 185))
        obj.draw(win)
        obj = graphics.Rectangle(graphics.Point(100, 165), graphics.Point(140, 185))
        obj.draw(win)
    
        # Текст "Вид носителя"
        text = graphics.Text(graphics.Point(60,175), "Вид носителя")
        text.setSize(10)
        text.draw(win)
    
    
        # Пишем вид носителя
        text = graphics.Text(graphics.Point(120,175), disk["Рассылка (общее кол-во дисков) (шт.):"])
        text.setSize(10)
        text.draw(win)        
    
    
        # Прямоуголник "Подразделение"
        obj = graphics.Rectangle(graphics.Point(308, 165), graphics.Point(428, 185))
        obj.draw(win)
        obj = graphics.Rectangle(graphics.Point(388, 165), graphics.Point(428, 185))
        obj.draw(win)
    
        # Текст "Подразделение"
        text = graphics.Text(graphics.Point(348,175), "Подразделение")
        text.setSize(10)
        text.draw(win)
    
        # Текст "0500"
        text = graphics.Text(graphics.Point(408,175), "0500")
        text.setSize(10)
        text.draw(win)
    
    
        # Прямоугольник "Контрольная характеристика"
        obj = graphics.Rectangle(graphics.Point(15, 215), graphics.Point(135, 295))
        obj.draw(win)
        obj = graphics.Rectangle(graphics.Point(15, 215), graphics.Point(135, 240))
        obj.draw(win)
    
        # Текст "Контрольная характеристика"
        text = graphics.Text(graphics.Point(75,230), "Контрольная характеристика")
        text.setSize(8)
        text.draw(win)
    
        # Текст c контрольной суммой из реферата
        # Разбиваем КСумм на две строки
        s = disk["Контрольная характеристика (сумма)"]
        l = len(s)
        s = '\n'.join([s[0:l//2], s[l//2::]])
    
        # Пишем Ксумм 
        text = graphics.Text(graphics.Point(75,265), s)
        text.setSize(10)
        text.draw(win)
    
    
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
        text = graphics.Text(graphics.Point(337,225), "ОТК")
        text.setSize(10)
        text.draw(win)    
        text = graphics.Text(graphics.Point(337,245), "Дата")
        text.setSize(10)
        text.draw(win)         
        text = graphics.Text(graphics.Point(337,265), "ВП МО")
        text.setSize(10)
        text.draw(win)     
        text = graphics.Text(graphics.Point(337,285), "Дата")
        text.setSize(10)
        text.draw(win) 


        # Прямоугольник для номера тома, рег. номера и вида ЭД
        obj = graphics.Rectangle(graphics.Point(78, 350), graphics.Point(370, 390))
        obj.draw(win)
        obj = graphics.Rectangle(graphics.Point(78, 350), graphics.Point(370, 370))
        obj.draw(win) 
        obj = graphics.Rectangle(graphics.Point(150, 350), graphics.Point(300, 390))
        obj.draw(win)

        # Текст "Номер тома/Количество томов"
        text = graphics.Text(graphics.Point(114,355), "Номер тома/")
        text.setSize(8)
        text.draw(win)
        text = graphics.Text(graphics.Point(114,365), "Количество томов")
        text.setSize(8)
        text.draw(win)

        # Пишем номер тома
        t = f"{disks.index(disk)+1}/{len(disks)}"
        text = graphics.Text(graphics.Point(114,380), t)
        text.setSize(8)
        text.draw(win)            

        # Текст "Регистрационный номер"
        text = graphics.Text(graphics.Point(225,360), "Регистрационный номер")
        text.setSize(8)
        text.draw(win)                 

        # Вставляем регистрационный номер из реферата
        text = graphics.Text(graphics.Point(225,380), disk["Регистрационный номер:"])
        text.setSize(8)
        text.draw(win) 

        # Текст "Вид ЭД"
        text = graphics.Text(graphics.Point(335,360), "Вид ЭД")
        text.setSize(8)
        text.draw(win)

        # Пишем, что формируем подлинник
        text = graphics.Text(graphics.Point(335,380), "П")
        text.setSize(8)
        text.draw(win)                            
            
        # Сохранение подлинника
        # Преобразуем изображение в формат postscript
        ps = win.postscript(colormode='color', pageheight=448, pagewidth=448)
        # win.postscript(colormode='color', file="image.eps")
        # Кодируем преобразованное изображение в байты
        img = Image.open(io.BytesIO(ps.encode('utf-8')))
        # Чтобы не потерять качество при преобразовании в .png увеличиваем
        # postscript изображение в 8 раз
        img.load(scale=8)
        # Возвращаем исходный масштаб
        img = img.resize((1394,1394), Image.BICUBIC)
        # Сохраняем изображение
        img.save("label" + str(disks.index(disk)) + "_p.png", "png", dpi=(300,300))
        
        # Ждем клика мышью на окне
        # win.getMouse()

        # Закрашиваем прямоугольник "Контрольная характеристика" для дубликата
        obj = graphics.Rectangle(graphics.Point(15, 215), graphics.Point(135, 295))
        obj.draw(win)
        obj.setFill("#87CEFA")
        obj = graphics.Rectangle(graphics.Point(15, 215), graphics.Point(135, 240))
        obj.draw(win)
    
        # Текст "Контрольная характеристика"
        text = graphics.Text(graphics.Point(75,230), "Контрольная характеристика")
        text.setSize(8)
        text.draw(win)

        # Закрашиваем прямоугольник "Вид ЭД"
        obj = graphics.Rectangle(graphics.Point(300, 370), graphics.Point(370, 390))
        obj.setFill("#87CEFA")
        obj.draw(win)

        # Пишем, что формируем дубликат
        text = graphics.Text(graphics.Point(335,380), "Д")
        text.setSize(8)
        text.draw(win)          

        # Ждем клика мышью на окне
        # win.getMouse()
        

        # Сохранение дубликата
        # Преобразуем изображение в формат postscript
        ps = win.postscript(colormode='color', pageheight=448, pagewidth=448)
        # win.postscript(colormode='color', file="image.eps")
        # Кодируем преобразованное изображение в байты
        img = Image.open(io.BytesIO(ps.encode('utf-8')))
        # Чтобы не потерять качество при преобразовании в .png увеличиваем
        # postscript изображение в 8 раз
        img.load(scale=8)
        # Возвращаем исходный масштаб
        img = img.resize((1394,1394), Image.BICUBIC)
        # Сохраняем изображение
        img.save("label" + str(disks.index(disk)) + "_d.png", "png", dpi=(300,300))
        
        #  Закрываем окно
        win.close()


# Проверяем, что скрипту передан в качестве аргумента файл реферата
if len(sys.argv) < 2:
        print("Необходимо указать файл реферата!")
        print("For help please input 'python label.py help'")
        sys.exit()
        
if sys.argv[1] == 'help':
    print("Usage: >python label.py 'referat filename' 'title field length'(default to 60) 'title text size'(default to 8)")
    print("Example: >python label.py ref.docx 40 10")
else:
    # Добавляем путь в SYSTEM PATH до Ghostscript под нужную архитектуру 
    path = os.path.dirname(os.path.abspath(sys.argv[0]))                # Абсолютный путь до папки со скриптом label.py
    
    # Определяем системную архитектуру и дополняем путь до папки с Ghostscript в зависимости от архитектуры
    if is_windows_64bit():
        path = path + r"\x64"
        
    else:
        path = path + r"\x32"
    
    # Модифицируем PATH
    app_path = os.path.join(path)
    os.environ["PATH"] += os.pathsep + app_path

    # Запуск основного скрипта
    main(*sys.argv[1:])

