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


# The base template for print
passportTemplate = DocxTemplate('wordDocuments' + sep + 'PassportTemplate.docx')
# The help template for create sticker in MarkSoft
helpTemplate = DocxTemplate('wordDocuments' + sep + 'HelpTemplate.docx')

context = {
    'basic_name': 'ВРУ-1-00-0-0-00-УХЛ4',
    'system_number': '1010',
    'name': 'ЩМП-1',
    'year': datetime.datetime.now().year,
    'nominal_current': 250,
    'nominal_Icu': 4.5,
    'IP': 54,
    'grounding': 'TN-C-S',
    'installation': 'навесной',
    'cross_section': '1x25',
    'height': 1000,
    'length': 800,
    'depth': 300,
    'mass': 50,
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


