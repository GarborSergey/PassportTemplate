import os
from os import sep
from docxtpl import DocxTemplate
import tkinter
from tkinter.ttk import Combobox, Checkbutton, Radiobutton, Progressbar
from tkinter import messagebox, filedialog, Menu
import datetime
import winshell

BASE_FONT = 'Arial Bold'
SIZE_BASE_FONT = 15
# Window size
WEIGHT = 800
HEIGHT = 500

# --------------------- WINDOW SETTINGS ---------------------
root = tkinter.Tk()
root.title('Auto Passport constructor')  # title window
root.iconbitmap(default='Pictures' + sep + 'AutoPassportIcon.ico')  # set app icon
root.attributes("-toolwindow", True)  # disable toolbar in window

root.geometry(f'{WEIGHT}x{HEIGHT}')
# -----------------------------------------------------------


# --------------- SET "ВРУ-X-XX-X-X-XX-УХЛ4" ----------------
# ------------------- LABEL AND COMBOBOX --------------------
# purposePanel
purposePanelLable = tkinter.Label(
    root,
    text='ВРУ-[X]-XX-X-X-XX-УХЛ4\nНазначение панели?',
    font=(BASE_FONT, SIZE_BASE_FONT)
)
purposePanelLable.grid(column=1, row=1)
purposePanelCombo = Combobox(root)
purposePanelCombo['values'] = (
    '1 - Вводное',
    '2 - Вводно-распределительное',
    '3 - Распределительное',
)
purposePanelCombo.grid(column=1, row=2)
# purposeIntroductionApparatus
purposeIntroductionApparatusLable = tkinter.Label(
    root,
    text='ВРУ-X-[XX]-X-X-XX-УХЛ4\nНазначение вводного аппарата управления?',
    font=(BASE_FONT, SIZE_BASE_FONT)
)
purposeIntroductionApparatusLable.grid(column=1, row=3)
purposeIntroductionApparatusCombo = Combobox(root)
purposeIntroductionApparatusCombo['values'] = (
    '0 - Отсутствует',
    '1 - Переключатель на 100А',
    '2 - Переключатель на 160А',
    '3 - Переключатель на 250А',
    '4 - Переключатель на 400А',
    '5 - Переключатель на 630А',
    '6 - Выключатель на 100А',
    '7 - Выключатель на 160А',
    '8 - Выключатель на 250А',
    '9 - Выключатель на 400А',
    '10 - Выключатель на 630А',
    '11 - Выключатель и аппаратура АВР на 100А',
    '12 - Выключатель и аппаратура АВР на 160А',
    '13 - Выключатель и аппаратура АВР на 250А',
    '14 - Выключатель и аппаратура АВР на 400А',
    '15 - Выключатель и аппаратура АВР на 630А',
    '16 - Два выключателя 100А',
    '17 - Два выключателя 160А',
    '18 - Два выключателя 250А',
    '19 - Два выключателя 400А',
    '20 - Два выключателя 630А',
    '21 - Выключатели и переключатели до 100А',
    '22 - Выключатели и переключатели более 630А',
    '23 - Выключатель и аппаратура АВР более 630А',
    '24 - Выключатель более 630А',
    '25 - Два выключателя боее 630А',
    '26 - Переключатель более 630А',
)
purposeIntroductionApparatusCombo.grid(column=1, row=4)  # FIX DESIGN MORE WEIGHT

existenceExtraDeviceLable = tkinter.Label(
    root,
    text='ВРУ-X-XX-[X]-X-XX-УХЛ4\nНаличие дополнительного оборудования',
    font=(BASE_FONT, SIZE_BASE_FONT)
)
# -----------------------------------------------------------


root.mainloop()

'''
# The base template for print
passportTemplate = DocxTemplate('wordDocuments' + sep + 'PassportTemplate.docx')
# The help template for create sticker in MarkSoft
helpTemplate = DocxTemplate('wordDocuments' + sep + 'HelpTemplate.docx')

ip = console_get_IP()

context = {
    'basic_name': console_get_basic_name(ip),
    'system_number': console_get_system_number(),
    'name': console_get_name(),
    'year': datetime.datetime.now().year,
    'nominal_current': console_get_nominal_current(),
    'nominal_Icu': console_get_nominal_Icu(),
    'IP': ip,
    'grounding': console_get_grounding(),
    'installation': console_get_installation(),
    'cross_section': console_get_cross_section(),
    'height': console_get_height(),
    'length': console_get_length(),
    'depth': console_get_depth(),
    'mass': console_get_mass(),
}

# Render templates
passportTemplate.render(context)
helpTemplate.render(context)

# Get path
fileName = f'passport_{context["name"]}_{context["system_number"]}'
passportPath = get_path_save_file(fileName)
helpPath = get_path_save_file(fileName + '_helper')

# Save files
passportTemplate.save(passportPath)
helpTemplate.save(helpPath)
'''