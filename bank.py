import openpyxl
from openpyxl.styles import Border, Side, Alignment

import argparse
import datetime
from glob import glob
import os
import logging
import pickle 

with open('sensetive.pkl', 'rb') as f:
    COMPANY = pickle.load(f)

number_format = '_-* #,##0.00\\ _₽_-;\\-* #,##0.00\\ _₽_-;_-* "-"??\\ _₽_-;_-@_-'
NUM_TAVRS = 3





def copy_bank(start_row, sheet_main, info):
    border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                    )

    for i, val in enumerate(info['DATA']):
        sheet_main[f'B{start_row + i}'].value = info['date'][i]
        sheet_main[f'B{start_row + i}'].number_format = 'dd/mm/yy;@'
        sheet_main[f'B{start_row + i}'].border = border

        sheet_main[f'C{start_row + i}'].value = val[0]
        sheet_main[f'C{start_row + i}'].number_format = number_format
        sheet_main[f'C{start_row + i}'].border = border

        sheet_main[f'D{start_row + i}'].value = val[1]
        sheet_main[f'D{start_row + i}'].number_format = number_format
        sheet_main[f'D{start_row + i}'].border = border

        sheet_main[f'E{start_row + i}'].value = val[2]
        sheet_main[f'E{start_row + i}'].number_format = '@'
        sheet_main[f'E{start_row + i}'].border = border
        sheet_main[f'E{start_row + i}'].alignment = Alignment(wrap_text = True)

        sheet_main[f'F{start_row + i}'].value = val[3]
        sheet_main[f'F{start_row + i}'].number_format = '@'
        sheet_main[f'F{start_row + i}'].border = border
        sheet_main[f'F{start_row + i}'].alignment = Alignment(wrap_text = True)

        sheet_main[f'G{start_row + i}'].border = border

        sheet_main[f'H{start_row + i}'].value = info['company']
        sheet_main[f'H{start_row + i}'].number_format = '@'
        sheet_main[f'H{start_row + i}'].border = border
    
    sheet_main[f'A{start_row + i}'].value = info['summ']
    sheet_main[f'A{start_row + i}'].number_format = '@'

    return start_row + i + 1





def get_sheet(path_to_main, path_to_tavr, name):
    # Читает и возвращает листы нужные для обработки
    main_xlsx = openpyxl.load_workbook(path_to_main)
    sheet_main = main_xlsx[name]

    sheet_from_list = [openpyxl.load_workbook(path, data_only=True)['Отчет 1'] for path in path_to_tavr]

    return main_xlsx, sheet_main, sheet_from_list





def get_info_bank(sheet_from):
    info = dict()

    try:
        key = int(sheet_from['A4'].value.split()[4])
    except:
        logging.error(f'Обратите внимание на ячейку A4, возможно она была изменена, если нет обратитесь к администратору')
        exit(2)

    info['company'] = COMPANY[key]
    info['date'] = []

    info['DATA'] = []
    num = 10
    while isinstance(sheet_from[f'A{num}'].value, datetime.datetime):

        info['date'].append(sheet_from[f'A{num}'].value)

        info['DATA'].append([sheet_from[f'G{num}'].value, # Debet
                             sheet_from[f'H{num}'].value, # Credit
                             sheet_from[f'I{num}'].value, # Назначение
                             sheet_from[f'J{num}'].value, # Контрагент
                             ])
        num+=1

    num+=1
    summ = sheet_from[f'A{num}'].value.split(':')[-1]
    info['summ'] = summ
    

    return info




def bank(path_to_main, path_to_bank, path_to_save):

    path_to_bank_list = glob(os.path.join(path_to_bank,'*.xlsx'))

    temp = len(path_to_bank_list)
    if temp:
        logging.info(f"Всего {temp} отчета найдено по банкам")
    else:
        logging.error("Не найдено ни одного отчета по банкам, проверьте павильность пути, если отчеты были загружены.")
        return

    main_xlsx, sheet_main, sheet_from_list = get_sheet(path_to_main, path_to_bank_list, name = 'БАНК')
    end = len(list(sheet_main.values)) + 1

    for i, sheet_from in enumerate(sheet_from_list):
        info = get_info_bank(sheet_from)
        start_row = end
        
        if len(info["DATA"]) == 0:
            logging.info(f"Отчет номер {i+1} пуст")
            continue

        end = copy_bank(start_row, sheet_main, info)
        logging.info(f'Отчет номер {i+1} добавлен')

    
    # Сохраняем обработанный файл main_xlsx
    main_xlsx.save(path_to_save)
    logging.info(f'Добавление и сохранение отчетов в банк завершено')