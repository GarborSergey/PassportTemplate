import os
from os import sep
from docxtpl import DocxTemplate
import datetime
import winshell


def get_path_save_file(filename: str):
    """
    Return path that need to save file in desktop, in dir Passport
    if dir not exists creating her.
    """

    desktop = winshell.desktop()

    if not os.path.exists(desktop + sep + 'Passports'):
        os.mkdir(desktop + sep + 'Passports')

    return desktop + sep + 'Passports' + sep + filename + '.docx'


def console_get_basic_name(ip):
    """Get console basic name"""
    # First param
    print('_'*50)
    print(f'ВРУ-[X]-XX-X-X-{ip}-УХЛ4'.center(50))
    print('_' * 50)
    print('Назначение панели?'.center(50))
    print('1 - Вводное')
    print('2 - Вводно-распределительное')
    print('3 - Распределительное')
    print('_' * 50)
    first = input('Введите соответствующий номер, нажмите ENTER: ')
    os.system('cls')

    # Second param
    print('_' * 50)
    print(f'ВРУ-{first}-[XX]-X-X-{ip}-УХЛ4'.center(50))
    print('_' * 50)
    print('Назначение вводного аппаратов управления?'.center(50))
    print('0 - Отсутствует')
    print('1 - Переключатель на 100А')
    print('2 - Переключатель на 160А')
    print('3 - Переключатель на 250А')
    print('4 - Переключатель на 400А')
    print('5 - Переключатель на 630А')
    print('6 - Выключатель на 100А')
    print('7 - Выключатель на 160А')
    print('8 - Выключатель на 250А')
    print('9 - Выключатель на 400А')
    print('10 - Выключатель на 630А')
    print('11 - Выключатель и аппаратура АВР на 100А')
    print('12 - Выключатель и аппаратура АВР на 160А')
    print('13 - Выключатель и аппаратура АВР на 250А')
    print('14 - Выключатель и аппаратура АВР на 400А')
    print('15 - Выключатель и аппаратура АВР на 630А')
    print('16 - Два выключателя 100А')
    print('17 - Два выключателя 160А')
    print('18 - Два выключателя 250А')
    print('19 - Два выключателя 400А')
    print('20 - Два выключателя 630А')
    print('21 - Выключатели и переключатели до 100А')
    print('22 - Выключатели и переключатели более 630А')
    print('23 - Выключатель и аппаратура АВР более 630А')
    print('24 - Выключатель более 630А')
    print('25 - Два выключателя боее 630А')
    print('26 - Переключатель более 630А')
    print('_' * 50)
    second = input('Введите соответствующий номер, нажмите ENTER: ')
    os.system('cls')

    # Third param
    print('_' * 50)
    print(f'ВРУ-{first}-{second}-[X]-X-{ip}-УХЛ4'.center(50))
    print('_' * 50)
    print('Наличие дополнительного оборудования?'.center(50))
    print('0 - Отсутствует')
    print('1 - Блок автоматического управления освещением на 30 групп')
    print('2 - Блок неавтоматического управления освещением на 30 групп')
    print('3 - Блок автоматического управления освещением на 14 групп')
    print('4 - Блок неавтоматического управления освещением на 14 групп')
    print('5 - Блок автоматического управления освещением на 8 групп')
    print('6 - Блок неавтоматического управления освещением на 8 групп')
    print('_' * 50)
    third = input('Введите соответствующий номер, нажмите ENTER: ')
    os.system('cls')

    # Fourth param
    print('_' * 50)
    print(f'ВРУ-{first}-{second}-{third}-[X]-{ip}-УХЛ4'.center(50))
    print('_' * 50)
    print('Защитные аппараты на отходящих линиях?'.center(50))
    print('1 - Автоматические выключатели')
    print('2 - Предохранители')
    fourth = input('Введите соответствующий номер, нажмите ENTER: ')
    if fourth == '1':
        fourth = 'A'
    else:
        fourth = None
    os.system('cls')

    if fourth:
        return f'ВРУ-{first}-{second}-{third}-{fourth}-{ip}-УХЛ4'
    else:
        return f'ВРУ-{first}-{second}-{third}-{ip}-УХЛ4'


def console_get_system_number():
    print('_' * 50)
    systemNumber = input('Введите номер щита в системе [0000], нажмите ENTER: ')
    os.system('cls')
    return systemNumber


def console_get_name():
    print('_' * 50)
    name = input('Введите название щита, нажмите ENTER: ')
    os.system('cls')
    return name


def console_get_nominal_current():
    print('_' * 50)
    nominalCurrent = input('Введите номинальный ток щита [A], нажмите ENTER: ')
    os.system('cls')
    return nominalCurrent


def console_get_nominal_Icu():
    print('_' * 50)
    nominalICU = input('Введите минимальный ток КЗ щита [kA], нажмите ENTER: ')
    os.system('cls')
    return nominalICU


def console_get_IP():
    print('_' * 50)
    ip = input('Введите степень защиты щита [IP00], нажмите ENTER: ')
    os.system('cls')
    return ip


def console_get_grounding():
    print('Система заземления шкафа'.center(50))
    print('1 - TN-C')
    print('2 - TN-S')
    print('3 - TN-C-S')
    print('_' * 50)
    num = int(input('Введите соответствующий номер, нажмите ENTER: '))
    grounding = {1: 'TN-C', 2: 'TN-S', 3: 'TN-C-S'}
    os.system('cls')
    return grounding[num]


def console_get_installation():
    print('Исполнение шкафа'.center(50))
    print('1 - Навесной')
    print('2 - Напольный')
    print('3 - Встраиваемый')
    print('_' * 50)
    num = int(input('Введите соответствующий номер, нажмите ENTER: '))
    installation = {1: 'навесной', 2: 'напольный', 3: 'встраиваемый'}
    os.system('cls')
    return installation[num]


def console_get_cross_section():
    print('_' * 50)
    crossSection = input('Максимальное сечение кабеля подкл. к вводному зажиму [0x00], нажмите ENTER: ')
    os.system('cls')
    return crossSection


def console_get_height():
    print('_' * 50)
    height = input('Введите высоту щита [мм], нажмите ENTER: ')
    os.system('cls')
    return height


def console_get_length():
    print('_' * 50)
    length = input('Введите ширину щита [мм], нажмите ENTER: ')
    os.system('cls')
    return length


def console_get_depth():
    print('_' * 50)
    depth = input('Введите глубину щита [мм], нажмите ENTER: ')
    os.system('cls')
    return depth


def console_get_mass():
    print('_' * 50)
    mass = input('Введите вес щита [кг], нажмите ENTER: ')
    os.system('cls')
    return mass


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
passportPath = get_path_save_file('new')
helpPath = get_path_save_file('new_helper')

# Save files
passportTemplate.save(passportPath)
helpTemplate.save(helpPath)


