import os
import re
import datetime
import tkinter

from os import sep
from docxtpl import DocxTemplate
from tkinter.ttk import Combobox
from tkinter import messagebox, filedialog, Checkbutton


VERSION = '2.5'  # 06/02/2023 13:19
BASE_FONT = 'Arial Bold'
SIZE_BASE_FONT = 13
BACKGROUND_COLOR = '#A0D6FF'
WINDOW_WEIGHT = 910
WINDOW_HEIGHT = 700
BASE_DIR = os.path.abspath(os.curdir)

# --------------------- WINDOW SETTINGS ---------------------
root = tkinter.Tk()
root.title('Auto Passport Constructor')  # title window
root.iconbitmap(BASE_DIR + sep + 'Pictures' + sep + 'AutoPassportIcon.ico')  # set app icon
root.configure(bg=BACKGROUND_COLOR)
root.geometry(f'{WINDOW_WEIGHT}x{WINDOW_HEIGHT}')
root.maxsize(WINDOW_WEIGHT, WINDOW_HEIGHT)
root.minsize(WINDOW_WEIGHT, WINDOW_HEIGHT)
# -----------------------------------------------------------


# ------------------------- FRAMES --------------------------
inputDataFrame = tkinter.LabelFrame(
    root,
    bd=0,
    bg=BACKGROUND_COLOR,
)

baseNameFrame = tkinter.LabelFrame(
    inputDataFrame,
    text='Условное обозначение',
    width=600,
    height=300,
    font=(BASE_FONT, SIZE_BASE_FONT + 2),
    bg=BACKGROUND_COLOR,
)

generalDataFrame = tkinter.LabelFrame(
    inputDataFrame,
    text='Общие данные',
    font=(BASE_FONT, SIZE_BASE_FONT + 2),
    width=600,
    height=300,
    bg=BACKGROUND_COLOR,
)

overallDimensionsWeightFrame = tkinter.LabelFrame(
    inputDataFrame,
    text='Габаритные размеры, вес',
    font=(BASE_FONT, SIZE_BASE_FONT + 2),
    width=600,
    height=300,
    bg=BACKGROUND_COLOR,
)

mainFrame = tkinter.LabelFrame(
    root,
    bd=0,
    bg=BACKGROUND_COLOR,
)

logoFrame = tkinter.LabelFrame(
    root,
    bd=0,
    width=600,
    height=300,
    bg=BACKGROUND_COLOR,
)
# -----------------------------------------------------------


# ----------------------- INPUT DATA ------------------------
# *************** SET "ВРУ-X-XX-X-X-XX-УХЛ4" ****************
purposePanelLable = tkinter.Label(
    baseNameFrame,
    text='Назначение панели: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
    bg=BACKGROUND_COLOR,
)
purposePanelCombo = Combobox(baseNameFrame, width=60, state="readonly")
purposePanelCombo['values'] = (
    '1 - Вводное',
    '2 - Вводно-распределительное',
    '3 - Распределительное',
)
# purposeIntroductionApparatus
purposeIntroductionApparatusLable = tkinter.Label(
    baseNameFrame,
    text='Назначение вводного аппарата управления: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
    bg=BACKGROUND_COLOR,
)
purposeIntroductionApparatusCombo = Combobox(baseNameFrame, width=60, state="readonly")
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
# existenceExtraDevice
existenceExtraDeviceLable = tkinter.Label(
    baseNameFrame,
    text='Наличие дополнительного оборудования: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
    bg=BACKGROUND_COLOR,
)
existenceExtraDeviceCombo = Combobox(baseNameFrame, width=60, state="readonly")
existenceExtraDeviceCombo['values'] = (
    '0 - Отсутствует',
    '1 - Блок автоматического управления освещением на 30 групп',
    '2 - Блок неавтоматического управления освещением на 30 групп',
    '3 - Блок автоматического управления освещением на 14 групп',
    '4 - Блок неавтоматического управления освещением на 14 групп',
    '5 - Блок автоматического управления освещением на 8 групп',
    '6 - Блок неавтоматического управления освещением на 8 групп',
)
# protectionOutgoingLines
protectionOutgoingLinesLable = tkinter.Label(
    baseNameFrame,
    text='Защитные аппараты на отходящих линиях: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
    bg=BACKGROUND_COLOR,
)
protectionOutgoingLinesCombo = Combobox(baseNameFrame, width=60, state="readonly")
protectionOutgoingLinesCombo['values'] = (
    '1 - Автоматические выключатели',
    '2 - Предохранители',
)
# internationalProtection
internationalProtectionLable = tkinter.Label(
    baseNameFrame,
    text='Степень защиты IP: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
    bg=BACKGROUND_COLOR,
)
internationalProtectionCombo = Combobox(baseNameFrame, width=60, state="readonly")
internationalProtectionCombo['values'] = (
    '30',
    '31',
    '54',
    '65',
)


# **********************************************************


# systemNumber
# ***************** Validate, Lable, Entry *****************
def is_valid_system_number(s):
    return re.match("^\d{0,5}$", s) is not None


check_system_number = (root.register(is_valid_system_number), "%P")

systemNumberLable = tkinter.Label(
    generalDataFrame,
    text='Номер заказа в базе: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
    bg=BACKGROUND_COLOR,
)
systemNumberEntry = tkinter.Entry(generalDataFrame, validate='key', validatecommand=check_system_number, width=15)


# **********************************************************


# nominalCurrent
# ***************** Validate, Lable, Entry *****************
def is_valid_nominal_current(s):
    return re.match("^\d{0,4}$", s) is not None


check_nominal_current = (root.register(is_valid_nominal_current), "%P")

nominalCurrentLable = tkinter.Label(
    generalDataFrame,
    text='Номинальный ток устройства [A]: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
    bg=BACKGROUND_COLOR,
)
nominalCurrentEntry = tkinter.Entry(generalDataFrame, validate='key', validatecommand=check_nominal_current, width=15)
# **********************************************************


# shortCircuitCurrent
# ********************** Lable, Entry **********************
shortCircuitCurrentLable = tkinter.Label(
    generalDataFrame,
    text='Минимальный ток КЗ [кА]:',
    font=(BASE_FONT, SIZE_BASE_FONT),
    bg=BACKGROUND_COLOR,
)
shortCircuitCurrentEntry = tkinter.Entry(generalDataFrame, width=15)
# **********************************************************


# grounding
# ******************** Lable, Combobox *********************
groundingLable = tkinter.Label(
    generalDataFrame,
    text='Система заземления: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
    bg=BACKGROUND_COLOR,
)
groundingCombo = Combobox(generalDataFrame, width=12, state="readonly")
groundingCombo['values'] = (
    'TN-S',
    'TN-C',
    'TN-C-S',
    'TT',
    'IT',
)
# **********************************************************


# crossSection
# ********************** Lable, Entry **********************
crossSectionLable = tkinter.Label(
    generalDataFrame,
    text='Максимальное сечение жил [Nxmm^2]: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
    bg=BACKGROUND_COLOR,
)
crossSectionEntry = tkinter.Entry(generalDataFrame, width=15)
# **********************************************************


# installation
# ******************** Lable, Combobox *********************
installationLable = tkinter.Label(
    generalDataFrame,
    text='Способ установки: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
    bg=BACKGROUND_COLOR,
)
installationCombo = Combobox(generalDataFrame, width=12, state="readonly")
installationCombo['values'] = (
    'Навесной',
    'Напольный',
    'Встраиваемый',
)


# **********************************************************


# height, length, depth, mass
# ***************** Validate, Lable, Entry *****************
def is_valid_HLDM(s):
    return re.match("^\d{0,5}$", s) is not None


check_HLDM = (root.register(is_valid_HLDM), "%P")

heightLable = tkinter.Label(
    overallDimensionsWeightFrame,
    text='Высота [мм]: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
    bg=BACKGROUND_COLOR,
)
lengthLable = tkinter.Label(
    overallDimensionsWeightFrame,
    text='Ширина [мм]: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
    bg=BACKGROUND_COLOR,
)
depthLable = tkinter.Label(
    overallDimensionsWeightFrame,
    text='Глубина [мм]: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
    bg=BACKGROUND_COLOR,
)
massLable = tkinter.Label(
    overallDimensionsWeightFrame,
    text='Вес [кг]: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
    bg=BACKGROUND_COLOR,
)

heightEntry = tkinter.Entry(overallDimensionsWeightFrame, validate='key', validatecommand=check_HLDM)
lengthEntry = tkinter.Entry(overallDimensionsWeightFrame, validate='key', validatecommand=check_HLDM)
depthEntry = tkinter.Entry(overallDimensionsWeightFrame, validate='key', validatecommand=check_HLDM)
massEntry = tkinter.Entry(overallDimensionsWeightFrame, validate='key', validatecommand=check_HLDM)
# **********************************************************


# name
# ********************** Lable, Entry **********************
nameLable = tkinter.Label(
    generalDataFrame,
    text='Название щита: ',
    font=(BASE_FONT, SIZE_BASE_FONT),
    bg=BACKGROUND_COLOR,
)
nameEntry = tkinter.Entry(generalDataFrame, width=15)
# **********************************************************


# createHelper
# **************** BooleanVar, Checkbutton *****************
createHelper = tkinter.BooleanVar()
createHelper.set(False)
createHelperCheckbutton = Checkbutton(
    mainFrame,
    text='Создать подсказку для шильдика',
    variable=createHelper,
    width=100,
    bg=BACKGROUND_COLOR,
    font=(BASE_FONT, SIZE_BASE_FONT),
    activebackground=BACKGROUND_COLOR,
    height=2
)


# **********************************************************
# -----------------------------------------------------------


# --------------- TRANSFORMATION INPUT DATA -----------------
# baseName
def construct_base_name_panel():
    try:
        purposePanel = purposePanelCombo.get()[0]
        purposeIntroductionApparatus = purposeIntroductionApparatusCombo.get()[:2]
        existenceExtraDevice = existenceExtraDeviceCombo.get()[0]
        internationalProtection = internationalProtectionCombo.get()[:2]

        if protectionOutgoingLinesCombo.get()[0] == '1':
            protectionOutgoingLines = True
        else:
            protectionOutgoingLines = False
    except IndexError:
        return None

    purposeIntroductionApparatus = purposeIntroductionApparatus.replace(' ', '')

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

# info function
def info():
    message = f'Auto Passport Constructor - version: {VERSION}\n' \
              f'creator: GarborSergey\n' \
              f'source: https://github.com/GarborSergey/PassportTemplate\n' \
              f'report bugs: garborfersru@gmail.com'
    messagebox.showinfo('Info', message)
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

    if not all([year, basicName, systemNumber, name, nominalCurrent, shortCircuitCurrent, internationalProtection,
                grounding, installation, crossSection, height, length, depth, mass]):
        return

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
    context = construct_context()

    if not context:
        messagebox.showerror('Error', 'Заполните все ячейки')
        return

    savePath = filedialog.askdirectory()

    fileName = construct_file_name()
    fileNameHelper = construct_file_name_helper()

    savePathFile = savePath + sep + fileName + '.docx'
    savePathFileHelper = savePath + sep + fileNameHelper + '.docx'

    passportTemplate = DocxTemplate('wordDocuments' + sep + 'PassportTemplate.docx')
    helpTemplate = DocxTemplate('wordDocuments' + sep + 'HelpTemplate.docx')

    passportTemplate.render(context)
    helpTemplate.render(context)

    passportTemplate.save(savePathFile)
    if createHelper.get():
        helpTemplate.save(savePathFileHelper)

    messagebox.showinfo('Success', f'"{fileName}" успешно создан по пути - [{savePath}]')
# -----------------------------------------------------------


# -------------- MAIN BUTTON CREATE FILE(S) -----------------
btn = tkinter.Button(
    mainFrame,
    text='СОЗДАТЬ ПАСПОРТ',
    command=create_file,
    font=(BASE_FONT, SIZE_BASE_FONT + 3)
)
# -----------------------------------------------------------


# ----------------------- POSITION --------------------------
logoFrame.grid(
    column=0,
    row=0,

    ipadx=10,
    ipady=10,
)

inputDataFrame.grid(
    column=0,
    row=1,

    ipadx=10,
    ipady=10,

    padx=10,
)
baseNameFrame.grid(
    column=0,
    row=1,

    columnspan=2,

    ipadx=10,
    ipady=10,

    padx=10,
)
generalDataFrame.grid(
    column=0,
    row=2,

    columnspan=1,

    ipadx=10,
    ipady=10,

    padx=10,

    sticky='w',
)
overallDimensionsWeightFrame.grid(
    column=1,
    row=2,

    columnspan=1,

    ipadx=10,
    ipady=10,

    padx=0,

    sticky='w',
)

mainFrame.grid(
    column=0,
    row=2,
)

# ********************** COMPANY LOGO **********************
logoCanvas = tkinter.Canvas(logoFrame, height=120, width=700, bg=BACKGROUND_COLOR, bd=0, highlightthickness=0)
logoImage = tkinter.PhotoImage(file=BASE_DIR + sep + 'Pictures' + sep + 'logo.png', )
logo = logoCanvas.create_image(0, 0, anchor='nw', image=logoImage)
logoCanvas.grid(column=0, row=0, columnspan=2)
# **********************************************************


# ********************** INFO BUTTON ***********************
infoImage = tkinter.PhotoImage(file=BASE_DIR + sep + 'Pictures' + sep + 'info.png')
smallInfoImage = infoImage.subsample(10, 10)
buttonInfo = tkinter.Button(
    logoFrame,
    image=smallInfoImage,
    bg=BACKGROUND_COLOR,
    borderwidth=0,
    activebackground=BACKGROUND_COLOR,
    command=info
)
buttonInfo.grid(column=3, row=0)
# **********************************************************

purposePanelLable.grid(column=0, row=1, sticky='e')
purposePanelCombo.grid(column=1, row=1)

purposeIntroductionApparatusLable.grid(column=0, row=2, sticky='e')
purposeIntroductionApparatusCombo.grid(column=1, row=2)

existenceExtraDeviceLable.grid(column=0, row=3, sticky='e')
existenceExtraDeviceCombo.grid(column=1, row=3)

protectionOutgoingLinesLable.grid(column=0, row=4, sticky='e')
protectionOutgoingLinesCombo.grid(column=1, row=4)

internationalProtectionLable.grid(column=0, row=5, sticky='e')
internationalProtectionCombo.grid(column=1, row=5)

nameLable.grid(column=0, row=0, sticky='e')
nameEntry.grid(column=1, row=0)

systemNumberLable.grid(column=0, row=1, sticky='e')
systemNumberEntry.grid(column=1, row=1)

nominalCurrentLable.grid(column=0, row=2, sticky='e')
nominalCurrentEntry.grid(column=1, row=2)

shortCircuitCurrentLable.grid(column=0, row=3, sticky='e')
shortCircuitCurrentEntry.grid(column=1, row=3)

groundingLable.grid(column=0, row=4, sticky='e')
groundingCombo.grid(column=1, row=4)

crossSectionLable.grid(column=0, row=5, sticky='e')
crossSectionEntry.grid(column=1, row=5)

installationLable.grid(column=0, row=6, sticky='e')
installationCombo.grid(column=1, row=6)

heightLable.grid(column=0, row=0, sticky='e')
heightEntry.grid(column=1, row=0)

lengthLable.grid(column=0, row=2, sticky='e')
lengthEntry.grid(column=1, row=2)

depthLable.grid(column=0, row=3, sticky='e')
depthEntry.grid(column=1, row=3)

massLable.grid(column=0, row=4, sticky='e')
massEntry.grid(column=1, row=4)

createHelperCheckbutton.configure()
createHelperCheckbutton.grid(column=0, row=0)

btn.configure(height=1, width=50, bg='#034569', fg='white')
btn.grid(column=0, row=1, pady=10, )

# -----------------------------------------------------------


# ---------------------CASE FOR TEST ------------------------
# purposePanelCombo.set('1')
# purposeIntroductionApparatusCombo.set('1')
# existenceExtraDeviceCombo.set('0')
# protectionOutgoingLinesCombo.set('1')
# internationalProtectionCombo.set('31')
# nameEntry.insert(0, 'TEST_ЩМП-1')
# systemNumberEntry.insert(0, '8888')
# nominalCurrentEntry.insert(0, '250')
# shortCircuitCurrentEntry.insert(0, '4.5')
# groundingCombo.set('TN-S')
# crossSectionEntry.insert(0, '1x25')
# installationCombo.set('Навесной')
# heightEntry.insert(0, '2000')
# lengthEntry.insert(0, '2000')
# depthEntry.insert(0, '2000')
# massEntry.insert(0, '2000')
# -----------------------------------------------------------


if __name__ == '__main__':
    root.mainloop()
