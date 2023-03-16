import pandas as pd
import numpy as np
import math
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import json


class Convertor:
    def __init__(self, excel_path):
        self.original = pd.read_excel(excel_path, 'исходный формат')
        self.original = self.original.fillna('')
        self.between = pd.DataFrame()
        data = {} # wtf bro
        self.result = pd.read_excel(excel_path, 'нужный формат')
        self.corr_fields = self.result.columns

    def dict_to_dataframe(self, items, df):
        if type(items) is dict:
            for key in items.keys():
                if type(items[key]) in [list, dict]:
                    self.dict_to_dataframe(items[key], df)
                else:
                    if df.__contains__(key):
                        if np.isnan(df[key].values(len(df) - 1)):
                            df.loc[len(df) - 1, key] = items[key]
                        else:
                            df.loc[len(df), key] = items[key]
                    else:
                        df[key] = [items[key]]
        elif type(items) is list:
            for item in items:
                self.dict_to_dataframe(item, df)

        return df

    def execute(self, command_data):
        """ Выполняет переданную команду, заполняя таблицы between и result
            args:
                command_data (tuple(str,str,str)): исполняемая команда \n
                1 аргумент- имя команды,\n
                2 аргумент - текущие колонки, \n
                3 аргумент - колонки которые должны получиться
        """
        command = command_data[0]
        input = command_data[1]
        corr = command_data[2]
        self.corr = corr

        changer = self.get_func(command)
        columns_for_change = [self.original[x] for x in input]
        corr_columns = changer(columns_for_change)

        for i in range(len(corr_columns[0]), self.original.shape[0]):
            for column in corr_columns:
                column.append(' ')

        for i in range(len(corr)):
            self.between[corr[i]] = corr_columns[i]
        self.fill_result()
        self.fix_date()

    def rename(self, columns_for_change):
        return columns_for_change

    def fill_result(self):
        result = pd.DataFrame()
        for field in self.corr_fields:
            if self.between.__contains__(field):
                result[field] = self.between[field]
            elif self.original.__contains__(field):
                result[field] = self.original[field]
            else:
                result[field] = [' ' for i in range(self.original.shape[0])]
        self.result = result

    def fix_date(self):
        for key in self.result.columns:
            if type(self.result[key][0]) is pd.Timestamp:
                list = []
                for item in self.result[key]:
                    a = str(item).replace('00:00:00', '')
                    list.append(a if a != 'NaT' else ' ')
                self.result[key] = list

    def get_func(self, command):
        if command == "SPLIT":
            return self.split_column
        elif command == "ZIP":
            return self.zip_columns
        elif command == "RENAME":
            return self.rename
        else:  # command == "RENAME":
            return self.empty_method

    # [["фамилия"],["имя"],["отчество"]] -> [["ФИО"]] #corr_columns[i]
    def split_column(self, columns_for_change):
        if type(columns_for_change[0][0]) in [pd.Timestamp, datetime]:
            return self.split_date(columns_for_change)
        splitter = ' '
        values = columns_for_change[0]
        length = len(self.corr)
        result = []
        for i in range(length):
            result.append([])
        for value in values:
            # проверка на длину
            if value == '':
                continue
            splitted_row = value.split(splitter)
            for i in range(length):
                result[i].append(splitted_row[i])
        return result

    def split_date(self, columns):
        values = columns[0]
        length = len(self.corr)
        result = []
        for i in range(length):
            result.append([])
        for value in values:
            if value == '':
                continue
            # проверка на длину
            date = [datetime.date(value), datetime.time(value)]
            # (value.year,value.month,value.day)
            for i in range(len(date)):
                result[i].append(date[i])
        return result

    def zip_date(self, columns_for_change):
        result = []
        for i in range(len(columns_for_change[0])):
            date = columns_for_change[0][i] if type(columns_for_change[0][i]) is pd.Timestamp else \
                columns_for_change[1][i]
            time = columns_for_change[0][i] if type(columns_for_change[0][i]) is not pd.Timestamp else \
                columns_for_change[1][i]
            if pd.isna(date) or pd.isna(time):
                result.append(' ')
                continue
            delta = timedelta(hours=time.hour, minutes=time.minute, seconds=time.second)
            result.append(date + delta)
        return [result]

    def zip_columns(self, columns_for_change):
        if type(columns_for_change[0][0]) in [pd.Timestamp, datetime]:
            return self.zip_date(columns_for_change)
        values = []
        for i in range(len(columns_for_change[0])):
            value = []
            for j in range(len(columns_for_change)):
                value.append(columns_for_change[j][i])
            values.append(' '.join(value))
        return [values]

    def have_empty_columns(self):
        for field in self.corr_fields:
            if self.between.__contains__(field):
                continue
            elif self.original.__contains__(field):
                continue
            else:
                return True
        return False

    def empty_method(self):
        pass

    def as_text(self, val):
        if val is None:
            return ""
        return str(val)

    def to_excel(self, path):
        """ Сохраняет xlsx файл с  таблицей result по указанному пути
            args:
                path (str): путь сохраняемого файла
        """
        writer = pd.ExcelWriter(path,
                                engine='openpyxl')
        self.result.to_excel(writer, 'нужный формат', index=False)

        wb = Workbook()
        ws = wb.active
        # добавляем строчки дфа в опенпайексель
        for r in dataframe_to_rows(self.result, index=False, header=True):
            ws.append(r)
        # серега что-то тут нашаманил
        for column in ws.columns:
            length = max(len(self.as_text(cell.value)) for cell in column)
            ws.column_dimensions[column[0].column_letter].width = length + 2
        wb.save(path)

    def to_json(self, path):
        """ Сохраняет json файл с  таблицей result по указанному пути
            args:
                path (str): путь сохраняемого файла
        """
        res = self.result.to_json(orient="index")
        parsed = json.loads(res)
        with open(path, 'w', encoding='utf-8') as outfile:
            json.dump(parsed, outfile, indent=4, ensure_ascii=False)

    def show_markdown(self):
        """ Формирует пердставление таблицы result в формате markdown
            :return:
                str: сформированная markdown таблица
        """
        # N = 3
        # shorted_res = self.result.iloc[:, 0:N]
        #
        # str_result = '| '
        # for field in shorted_res:
        #     str_result += f'{str(field)} |'
        # str_result += '\n| '
        # for i in range(len(shorted_res)):
        #     str_result += '--- | '
        # str_result += '\n| '
        # for i, row in shorted_res.iterrows():
        #     for field in row:
        #         str_result += f'{str(field)} |'
        #     str_result += '\n| '
        # return str_result[:-2]

        # for all
        result = '| '
        for field in self.result.columns:
            result += f'{str(field)} |'
        result += '\n| '
        for i in range(len(self.result.columns)):
            result += '--- | '
        result += '\n| '
        for i, row in self.result.iterrows():
            for field in row:
                result += f'{str(field)} |'
            result += '\n| '
        return result[:-2]

    def to_markdown(self, path):
        """ Сохраняет текстовый файл с markdown таблицей result по указанному пути
            args:
                path (str): путь сохраняемого файла
        """
        with open(path, "w") as file:
            file.write(self.show_markdown())