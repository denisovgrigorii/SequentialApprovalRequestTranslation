import json
import openpyxl
import zipfile
from openpyxl import Workbook
from openpyxl.styles import Font

#чек
class SequentialApprovalRequestTranslation:
    def __init__(self):
        self.wb = Workbook()
        self.ws = self.wb.active
        self.raw_in_column = 2
        self.separator = '\u2192'  # красивый разделить для excel между ролями
        self.mass_for_pars = []
        self.dict_1 = {}  # fieldName + templates.managed.form.organization\templates.managed.form.is.
        self.dict_2 = {}  # зависимость имя роли в stages -fieldName
        self.number_role_field_name = 0
        # cловарь декодироввных переменных и их fieldName
        self.result_dict_value = {}
        self.dict_3 = {}
        with open('SequentialApprovalRequest.json', 'r', encoding='utf-8') as json_file:
            self.sequential_approval_request = json.load(json_file)  # cериализация json файла
        self.name_ir_list = list(self.sequential_approval_request.keys())
        self.number_name_ir_list = 0
        self.name_role_for_ir = list(self.sequential_approval_request
                                     [self.name_ir_list[self.number_name_ir_list]]['roleVariables'].keys())
        self.raw_in_column = 2

    def work_with_excel_file(self):
        # заголовки столбцов + выделение жирным шрифтом
        self.ws['A1'] = 'Имя информационного ресурса'
        self.ws['B1'] = 'Этапы согласования'
        self.ws['A1'].font = Font(bold=True)
        self.ws['B1'].font = Font(bold=True)
        # размеры слобцов
        self.ws.column_dimensions['A'].width = 50
        self.ws.column_dimensions['B'].width = 120
        for i in self.sequential_approval_request.keys():
            self.number_column_a = 'A{}'.format(self.raw_in_column)
            self.number_column_b = 'B{}'.format(self.raw_in_column)
            self.row_role = "".join(str(self.sequential_approval_request
                                        [self.name_ir_list[self.number_name_ir_list]]['stages'])
                                    .replace("[", "").replace("]", "")).replace("'", "") .replace(' $', '$') .split(",")

    def pars_lookup_name_variables(self):
        for role in self.row_role:
            if '$' in role:
                self.mass_for_pars.append(role)
        # замена имени роли на fieldName
        for role_name in self.mass_for_pars:
            self.field_name_role = \
                self.sequential_approval_request[i]['roleVariables'][self.role_name.replace(' ', '')]['fieldName']
            self.managed_object = \
                self.sequential_approval_request[i]['roleVariables'][self.role_name.replace(' ', '')]['managedObject']
            dict_2[role_name] = field_name_role
            if managed_object == 'managed/is':
                self.mass_for_pars[s] = str('templates.managed.form.is.' + field_name_role)
                self.dict_1[field_name_role] = mass_for_pars[s]
            elif managed_object == 'managed/organization':
                self.mass_for_pars[s] = str('templates.managed.form.organization.' + field_name_role)
                self.dict_1[field_name_role] = self.mass_for_pars[s]
            s += 1
            return self.mass_for_pars

    # поиск значений переменных по файлам translation_ru.properties
    def find_value(name_value_for_pars: list, url_translation_ru: str = 'translation_ru.properties') -> list:
        dict_value = {}
        if len(name_value_for_pars) > 0:
            with open(url_translation_ru, 'r') as dict_file:
                for row in dict_file:
                    if len(dict_value) < len(name_value_for_pars):
                        # если какое либо значение из массива есть в строке - записать в find_value
                        find_value = [x for x in name_value_for_pars if x in row]
                        if find_value:
                            # удаляем все символы до знака "="
                            dict_value[find_value[0]] = row.split('=', 1)[1].lstrip()
                dict_file.close()
            with open('translation_r.properties', 'r') as dict_file:
                for row in dict_file:
                    if len(dict_value) < len(name_value_for_pars):
                        find_value = [x for x in name_value_for_pars if x in row]
                        if find_value:
                            # удаляем все символы до знака "="
                            dict_value[find_value[0]] = row.split('=', 1)[1].lstrip()
                dict_file.close()
        return dict_value

    def decode(unicode_role: str) -> str:
        return unicode_role.encode().decode('unicode-escape')

    def main_f(self):
        self.work_with_excel_file()
        self.pars_lookup_name_variables()
        # отправляем парсить и получать значения в функцию
        if len(self.mass_for_pars) > 0:
            return_mass_value = find_value(self.mass_for_pars)
        # decode полученных переменных
        for value in return_mass_value:
            decode_value = decode(return_mass_value[value])
            self.result_dict_value[value] = decode_value
            for value in result_dict_value.keys():
                if value in dict_1.values() and 'templates.managed.form.organization.' in value:
                    dict_3[str(value).replace('templates.managed.form.organization.', '')] = result_dict_value[value]
                elif value in dict_1.values() and 'templates.managed.form.is.' in value:
                    dict_3[str(value).replace('templates.managed.form.is.', '')] = result_dict_value[value]
            m = 0
            value_result_list = [value for value in dict_3.values()]
            for value in dict_2.keys():
                dict_2[value] = value_result_list[m]
                m += 1
            g = 0
            # удаляем возможные пробелы в ключах словаря(костыль)
            new_dict_2 = {}
            for k, v in dict_2.items():
                new_dict_2[k.strip()] = v
            # обновляем list stages c ролями для ИС
            self.row_role = [row.strip(' ') for row in row_role]
            for row in row_role:
                if row in new_dict_2.keys():
                    row_role[g] = new_dict_2[row]
                g += 1
            # проверяем есть ли линейный - если есть добавляем его в стоблец В
            if sequential_approval_request[name_ir_list[number_name_ir_list]]['managerStage']['isEnabled'] == True \
                    and len(row_role) != 0:
                row_role.insert(0, 'Линейный руководитель')

                if row_role[1] == '':  # проверка что в массиве нет пустых значений
                    row_role.pop(1)
                ws['{}'.format(number_column_b)] = '{}' \
                    .format(separator.join(row_role)).replace('[', '').replace(']', '').replace("'", "").replace('\\n',
                                                                                                                 '')
            else:
                ws['{}'.format(number_column_b)] = '{}' \
                    .format(separator.join(row_role)).replace('[', '').replace(']', '').replace("'", "").replace('\\n',
                                                                                                                 '')
            raw_in_column += 1
            number_name_ir_list += 1
            if len(name_role_for_ir) < number_role_field_name:
                number_role_field_name += 1
        self.wb.save("Этапы согласования.xlsx")


if __name__ == '__main__':
    start = SequentialApprovalRequestTranslation()
    start.main_f()