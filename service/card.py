import pythoncom
import win32com.client
import os

from data.utils import get_file, get_order
#
# name_temp = "Temp1"
# name = "Project4"
#
# # Links on Template
# temp_dir = "D:\.cardex\card_templates" + chr(92) + name_temp + chr(92)
# file_list = os.listdir(temp_dir)
# temp_file_front = name_temp + " (1).ai"
# temp_file_back = name_temp + " (2).ai"
#
# # Links on Project
# project_dir = "D:\.cardex\card" + chr(92) + name + chr(92)
# file_front = name + " (1).ai"
# file_back = name + " (2).ai"
# file_front_jpg = name + " (1).jpg"
# file_back_jpg = name + " (2).jpg"
#
# # Command for opening 'Template front file' with Terminal
# open_front = "Illustrator " + chr(34) + temp_dir + front_temp + chr(34)
# # Command for opening 'Template back file' with Terminal
# open_back = "Illustrator " + chr(34) + templ_dir + templ_file_back + chr(34)

# Running Illustrator
ai = win32com.client.Dispatch("Illustrator.Application")


def fn(dp, order_id):

    # Export Options for JPG
    exportOptions = win32com.client.Dispatch("Illustrator.ExportOptionsJPEG")
    exportOptions.qualitySetting = 100
    type_ = 1

    order = get_order(order_id)


    front_temp = get_file(order.get("front"), "ai")
    os.system(front_temp)
    front = dp.ActiveDocument
    card = front.Layers("Визитка")
    
    # name = card.TextFrames("Имя")
    # num1 = card.TextFrames("Номер 1")
    # num2 = card.TextFrames("Номер 2")
    # num3 = card.TextFrames("Номер 3")
    # address  = card.TextFrames("Адрес")

    # name.TextRange.Contents = "Имя"
    # num1.TextRange.Contents = "(00) 000-00-00"
    # num2.TextRange.Contents = "(11) 111-11-11"
    # num3.TextRange.Contents = "(22) 222-22-22"
    # address.TextRange.Contents = "Адрес"

    try:
        location = card.TextFrames("Местоположение")
        location.TextRange.Contents = "Местоположение"
    except pythoncom.com_error as error:
        pass

    try:
        social = card.TextFrames("Соц-Сеть")
        social.TextRange.Contents = "@username"
    except pythoncom.com_error as error:
        pass

    dp.activeDocument.Export(project_dir + file_front_jpg, type_, exportOptions)
    front.SaveAs(project_dir + file_front)
    front.Close()

    back_temp = get_file(order.get("back"), "ai")
    
    os.system(back_temp)
    back = dp.ActiveDocument
    card = back.Layers("Визитка")
    
    try:
        social = card.TextFrames("Соц-Сеть")
        social.TextRange.Contents = "@username"
    except pythoncom.com_error as error:
        pass

    dp.activeDocument.Export(project_dir + file_back_jpg, type_, exportOptions)
    back.SaveAs(project_dir + file_back)
    back.Close()
    
    os.system("python -u D:/.cardex/scripts/mockup.py")
