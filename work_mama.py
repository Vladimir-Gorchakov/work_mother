import openpyxl
from openpyxl.styles import Border, Side, Alignment

import argparse
import datetime
from glob import glob
import os
import logging

from cash import new_cash
from bank import bank

number_format = '_-* #,##0.00\\ _₽_-;\\-* #,##0.00\\ _₽_-;_-* "-"??\\ _₽_-;_-@_-'
NUM_TAVRS = 3



def setup_parser()-> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()

    parser.add_argument("-b", "--bank", 
                        action="store_true", 
                        help="Добавить в банк")
    
    parser.add_argument("-nc","--new_cash", 
                        action="store_true", 
                        help="Добавить в новую кассу")
    
    parser.add_argument("--color", 
                        default = 'FF98DF17', 
                        type = str, 
                        help="Выбрать цвет для раздения дней тут https://ankiewicz.com/color/picker")
    return parser






def main(args):
    # Путь к главному файлу, куда добавляем отчеты
    path_to_main =  r'C:\Users\Пользователь\Desktop\WORK\КАССА_БАНК_2024.xlsx'
    # Заменить старый отчет новым - True, Сделать новый файл - False
    # Не рекомендуется
    overwrite = True
    # Путь к папке с отчетами
    path_to_tavr = r'C:\Users\Пользователь\Desktop\WORK\отчеты'
    # Путь к папке с отчетами банка
    path_to_bank = r'C:\Users\Пользователь\Desktop\WORK\отчеты_банк'
    # Путь для сохранения файла
    #path_to_save = "/home/linux228/finance/output.xlsx"

    if overwrite:
        path_to_save = path_to_main

    if args.new_cash:
        new_cash(path_to_main, path_to_tavr, path_to_save, args.color)

    if args.bank:
        bank(path_to_main, path_to_bank, path_to_save)




if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO,
                        format='%(levelname)s : %(message)s')
    parser = setup_parser()

    args = parser.parse_args()
    main(args)