import os
from os import sep
from docxtpl import DocxTemplate
import tkinter
from tkinter.ttk import Combobox, Checkbutton, Radiobutton, Progressbar
from tkinter import messagebox, filedialog, Menu
import datetime
import winshell
import re

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


# ----------------------- INPUT DATA ------------------------
# *************** SET "ВРУ-X-XX-X-X-XX-УХЛ4" ****************
purposePanelLable = tkinter.Label(
    root,
    text='ВРУ-[X]-XX-X-X-XX-УХЛ4\nНазначение панели?',
    font=(BASE_FONT, SIZE_BASE_FONT),
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
    font=(BASE_FONT, SIZE_BASE_FONT),
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
# existenceExtraDevice
existenceExtraDeviceLable = tkinter.Label(
    root,
    text='ВРУ-X-XX-[X]-X-XX-УХЛ4\nНаличие дополнительного оборудования',
    font=(BASE_FONT, SIZE_BASE_FONT),
)
existenceExtraDeviceLable.grid(column=1, row=5)
existenceExtraDeviceCombo = Combobox(root)
existenceExtraDeviceCombo['values'] = (
    '0 - Отсутствует',
    '1 - Блок автоматического управления освещением на 30 групп',
    '2 - Блок неавтоматического управления освещением на 30 групп',
    '3 - Блок автоматического управления освещением на 14 групп',
    '4 - Блок неавтоматического управления освещением на 14 групп',
    '5 - Блок автоматического управления освещением на 8 групп',
    '6 - Блок неавтоматического управления освещением на 8 групп',
)
existenceExtraDeviceCombo.grid(column=1, row=6)
# protectionOutgoingLines
protectionOutgoingLinesLable = tkinter.Label(
    root,
    text='ВРУ-X-XX-X-[X]-XX-УХЛ4\nЗащитные аппараты на отходящих линиях',
    font=(BASE_FONT, SIZE_BASE_FONT),
)
protectionOutgoingLinesLable.grid(column=1, row=7)
protectionOutgoingLinesCombo = Combobox(root)
protectionOutgoingLinesCombo['values'] = (
    '1 - Автоматические выключатели',
    '2 - Предохранители',
)
protectionOutgoingLinesCombo.grid(column=1, row=8)
# internationalProtection
internationalProtectionLable = tkinter.Label(
    root,
    text='ВРУ-X-XX-X-X-[XX]-УХЛ4\nСтепень защиты IP',
    font=(BASE_FONT, SIZE_BASE_FONT),
)
internationalProtectionLable.grid(column=1, row=9)
internationalProtectionCombo = Combobox(root)
internationalProtectionCombo['values'] = (
    '30',
    '31',
    '54',
    '65',
)
internationalProtectionCombo.grid(column=1, row=10)
# **********************************************************


# systemNumber
# ***************** Validate, Lable, Entry *****************
def is_valid_system_number(s):
    return re.match("^\d{0,5}$", s) is not None


check_system_number = (root.register(is_valid_system_number), "%P")

systemNumberLable = tkinter.Label(
    root,
    text='Номер заказа в базе: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
)
systemNumberLable.grid(column=2, row=1)
systemNumberEntry = tkinter.Entry(validate='key', validatecommand=check_system_number)
systemNumberEntry.grid(column=3, row=1)
# **********************************************************


# nominalCurrent
# ***************** Validate, Lable, Entry *****************
def is_valid_nominal_current(s):
    return re.match("^\d{0,4}$", s) is not None


check_nominal_current = (root.register(is_valid_nominal_current), "%P")

nominalCurrentLable = tkinter.Label(
    root,
    text='Номинальный ток устройства: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
)
nominalCurrentLable.grid(column=2, row=2)
nominalCurrentEntry = tkinter.Entry(validate='key', validatecommand=check_nominal_current)
nominalCurrentEntry.grid(column=3, row=2)
# **********************************************************


def construct_base_name_panel():
    purposePanel = purposePanelCombo.get()[0]
    purposeIntroductionApparatus = purposeIntroductionApparatusCombo.get()[:2]
    purposeIntroductionApparatus = purposeIntroductionApparatus.replace(' ', '')
    existenceExtraDevice = existenceExtraDeviceCombo.get()[0]

    if protectionOutgoingLinesCombo.get()[0] == '1':
        protectionOutgoingLines = True
    else:
        protectionOutgoingLines = False

    internationalProtection = internationalProtectionCombo.get()[:2]

    if protectionOutgoingLines:
        base_name = f'ВРУ-{purposePanel}-{purposeIntroductionApparatus}-' \
                    f'{existenceExtraDevice}-A-{internationalProtection}-УХЛ4'
    else:
        base_name = f'ВРУ-{purposePanel}-{purposeIntroductionApparatus}-' \
                    f'{existenceExtraDevice}-{internationalProtection}-УХЛ4'

    return base_name
# -----------------------------------------------------------


# Choose save path to directory without filename FIX FILENAME
def get_save_path():
    save_path = filedialog.askdirectory()
    return save_path


get_save_path_button = tkinter.Button(root, text='Save in', command=get_save_path)
get_save_path_button.grid(column=1, row=12)

btn = tkinter.Button(root, text='TEST', command=construct_base_name_panel)
btn.grid(column=1, row=11)


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