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
WEIGHT = 1000
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
    text='Номинальный ток устройства [A]: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
)
nominalCurrentLable.grid(column=2, row=2)
nominalCurrentEntry = tkinter.Entry(validate='key', validatecommand=check_nominal_current)
nominalCurrentEntry.grid(column=3, row=2)
# **********************************************************


#shortCircuitCurrent
# ********************** Lable, Entry **********************
shortCircuitCurrentLable = tkinter.Label(
    root,
    text='Минимальный ток КЗ [кА]:',
    font=(BASE_FONT, SIZE_BASE_FONT),
)
shortCircuitCurrentLable.grid(column=2, row=3)
shortCircuitCurrentEntry = tkinter.Entry()
shortCircuitCurrentEntry.grid(column=3, row=3)
# **********************************************************


#grounding
# ******************** Lable, Combobox *********************
groundingLable = tkinter.Label(
    root,
    text='Система заземления: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
)
groundingLable.grid(column=2, row=4)
groundingCombo = Combobox(root)
groundingCombo['values'] = (
    'TN-S',
    'TN-C',
    'TN-C-S',
    'TT',
    'IT',
)
groundingCombo.grid(column=3, row=4)
# **********************************************************


#crossSection
# ********************** Lable, Entry **********************
crossSectionLable = tkinter.Label(
    root,
    text='Максимальное сечение жил [Nxmm^2]: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
)
crossSectionLable.grid(column=2, row=5)
crossSectionEntry = tkinter.Entry()
crossSectionEntry.grid(column=3, row=5)
# **********************************************************


#installation
# ******************** Lable, Combobox *********************
installationLable = tkinter.Label(
    root,
    text='Способ установки: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
)
installationLable.grid(column=2, row=6)
installationCombo = Combobox(root)
installationCombo['values'] = (
    'Навесной',
    'Напольный',
    'Встраиваемый',
)
installationCombo.grid(column=3, row=6)
# **********************************************************


# height, length, depth, mass
# ***************** Validate, Lable, Entry *****************
def is_valid_HLDM(s):
    return re.match("^\d{0,5}$", s) is not None


check_HLDM = (root.register(is_valid_HLDM), "%P")

heightLable = tkinter.Label(
    root,
    text='Высота [мм]: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
)
lengthLable = tkinter.Label(
    root,
    text='Ширина [мм]: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
)
depthLable = tkinter.Label(
    root,
    text='Глубина [мм]: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
)
massLable = tkinter.Label(
    root,
    text='Вес [кг]: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
)
heightLable.grid(column=2, row=7)
lengthLable.grid(column=2, row=8)
depthLable.grid(column=2, row=9)
massLable.grid(column=2, row=10)

heightEntry = tkinter.Entry(validate='key', validatecommand=check_HLDM)
lengthEntry = tkinter.Entry(validate='key', validatecommand=check_HLDM)
depthEntry = tkinter.Entry(validate='key', validatecommand=check_HLDM)
massEntry = tkinter.Entry(validate='key', validatecommand=check_HLDM)

heightEntry.grid(column=3, row=7)
lengthEntry.grid(column=3, row=8)
depthEntry.grid(column=3, row=9)
massEntry.grid(column=3, row=10)
# **********************************************************


# name
# ********************** Lable, Entry **********************
nameLable = tkinter.Label(
    root,
    text='Название щита: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
)
nameLable.grid(column=2, row=11)
nameEntry = tkinter.Entry()
nameEntry.grid(column=3, row=11)
# **********************************************************


# createHelper
# **************** BooleanVar, Checkbutton *****************
createHelper = tkinter.BooleanVar()
createHelper.set(False)
createHelperCheckbutton = Checkbutton(
    root,
    text='Создать подсказку для шильдика',
    variable=createHelper,
    state='DISABLE',
)
createHelperCheckbutton.grid(column=3, row=12)
# **********************************************************
# -----------------------------------------------------------


# --------------- TRANSFORMATION INPUT DATA -----------------
# baseName
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


# fileName
def construct_file_name():
    name = nameEntry.get()
    systemNumber = systemNumberEntry.get()
    return f'{name}_{systemNumber}_passport'


# fileNameHelper
def construct_file_name_helper():
    name = nameEntry.get()
    systemNumber = systemNumberEntry.get()
    return f'{name}_{systemNumber}_helper'
# -----------------------------------------------------------


# ------------------- CONSTRUCT CONTEXT ---------------------
def construct_context():
    year = datetime.datetime.now().year
    basicName = construct_base_name_panel()
    systemNumber = systemNumberEntry.get()
    name = nameEntry.get()
    nominalCurrent = nominalCurrentEntry.get()
    shortCircuitCurrent = shortCircuitCurrentEntry.get()
    internationalProtection = internationalProtectionCombo.get()
    grounding = groundingCombo.get()
    installation = installationCombo.get()
    crossSection = crossSectionEntry.get()
    height = heightEntry.get()
    length = lengthEntry.get()
    depth = depthEntry.get()
    mass = massEntry.get()

    context = {
        'basic_name': basicName,
        'system_number': systemNumber,
        'name': name,
        'year': year,
        'nominal_current': nominalCurrent,
        'nominal_Icu': shortCircuitCurrent,
        'IP': internationalProtection,
        'grounding': grounding,
        'installation': installation,
        'cross_section': crossSection,
        'height': height,
        'length': length,
        'depth': depth,
        'mass': mass,
    }

    return context
# -----------------------------------------------------------


# ------------- MAIN FUNCTION CREATE FILE(S) ----------------
def create_file():
    savePath = filedialog.askdirectory()

    fileName = construct_file_name()
    fileNameHelper = construct_file_name_helper()

    savePathFile = savePath + sep + fileName + '.docx'  # MB EXCEPTION
    savePathFileHelper = savePath + sep + fileNameHelper + '.docx'

    context = construct_context()

    passportTemplate = DocxTemplate('wordDocuments' + sep + 'PassportTemplate.docx')
    helpTemplate = DocxTemplate('wordDocuments' + sep + 'HelpTemplate.docx')

    passportTemplate.render(context)
    helpTemplate.render(context)

    passportTemplate.save(savePathFile)
    if createHelper:
        helpTemplate.save(savePathFileHelper)

# -----------------------------------------------------------


# -------------- MAIN BUTTON CREATE FILE(S) -----------------
btn = tkinter.Button(
    root,
    text='Сохранить',
    command=create_file,
)
btn.grid(column=3, row=13)
# -----------------------------------------------------------

if __name__ == '__main__':
    root.mainloop()

