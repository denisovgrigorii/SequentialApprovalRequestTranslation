import json
import zipfile
import paramiko
import getpass
import os
import shutil
from openpyxl import Workbook
from openpyxl.styles import Font


def main_func():
    wb = Workbook()
    ws = wb.active
    separator = '\u2192'  # красивый разделить для excel между ролями
    with open('SequentialApprovalRequest.json', 'r', encoding='utf-8') as json_file:
        sequential_approval_request = json.load(json_file)  # cериализация json файла
    # заголовки столбцов + выделение жирным шрифтом
    ws['A1'] = 'Имя информационного ресурса'
    ws['B1'] = 'Этапы согласования'
    ws['A1'].font = Font(bold=True)
    ws['B1'].font = Font(bold=True)
    # размеры слобцов
    ws.column_dimensions['A'].width = 50
    ws.column_dimensions['B'].width = 120
    # переменные для работы с cериализованным json
    name_ir_list = list(sequential_approval_request.keys())
    number_name_ir_list = 0
    raw_in_column = 2
    for i in sequential_approval_request.keys():
        number_column_a = 'A{}'.format(raw_in_column)
        number_column_b = 'B{}'.format(raw_in_column)
        ws['{}'.format(number_column_a)] = i
        # собираем все роли для одной ИР в строку
        row_role = "".join(str(sequential_approval_request[name_ir_list[number_name_ir_list]]['stages'])
                           .replace("[", "").replace("]", "")).replace("'", "").replace(' $', '$').split(",")
        s = 0
        # парсим lookup значения переменных
        mass_for_pars = []
        for role in row_role:
            if '$' in role:
                mass_for_pars.append(role)
        # fieldName + templates.managed.form.organization\templates.managed.form.is.
        dict_1 = {}
        # зависимость имя роли в stages -fieldName
        dict_2 = {}
        # замена имени роли на fieldName
        for role_name in mass_for_pars:
            field_name_role = sequential_approval_request[i]['roleVariables'][role_name.replace(' ', '')]['fieldName']
            managed_object = sequential_approval_request[i]['roleVariables']\
                                                        [role_name.replace(' ', '')]['managedObject']
            dict_2[role_name] = field_name_role
            if managed_object == 'managed/is':
                mass_for_pars[s] = str('templates.managed.form.is.' + field_name_role)
                dict_1[field_name_role] = mass_for_pars[s]
            elif managed_object == 'managed/organization':
                mass_for_pars[s] = str('templates.managed.form.organization.' + field_name_role)
                dict_1[field_name_role] = mass_for_pars[s]
            s += 1
        url_translation_ru = jar_unzip() # если получится реализовать запуск на linux
        # отправляем парсить и получать значения в функцию
        if len(mass_for_pars) > 0:
            return_mass_value = find_value(mass_for_pars,url_translation_ru)
        # cловарь декодироввных переменных и их fieldName
            result_dict_value = {}
            dict_3 = {}
            # decode полученных переменных
            for value in return_mass_value:
                decode_value = decode(return_mass_value[value])
                result_dict_value[value] = decode_value
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
            row_role = [row.strip(' ') for row in row_role]
            for row in row_role:
                if row in new_dict_2.keys():
                    row_role[g] = new_dict_2[row]
                g +=1
        # проверяем есть ли линейный - если есть добавляем его в стоблец В
        try:
            if sequential_approval_request[name_ir_list[number_name_ir_list]]['managerStage']['isEnabled'] == True:
                row_role.insert(0, 'Линейный руководитель')
                if row_role[1] == '':  # проверка что в массиве нет пустых значений
                    row_role.pop(1)
                ws['{}'.format(number_column_b)] = '{}' \
                    .format(separator.join(row_role)).replace('[', '')\
                    .replace(']', '')\
                    .replace("'", "")\
                    .replace('\\n', '')
            else:
                ws['{}'.format(number_column_b)] = '{}'\
                    .format(separator.join(row_role)).replace('[', '')\
                    .replace(']', '')\
                    .replace("'", "")\
                    .replace('\\n', '')
        except KeyError:
            ws['{}'.format(number_column_b)] = '{}' \
                .format(separator.join(row_role))\
                .replace('[', '').replace(']', '')\
                .replace("'", "")\
                .replace('\\n', '')
        raw_in_column += 1
        number_name_ir_list += 1
    wb.save("Этапы согласования.xlsx")
    print("Файл Этапы согласования.xlsx успешно создан в текущей директории")


# поиск значений переменных по файлам translation_ru.properties
def find_value(name_value_for_pars: list = [], url_translation_ru: str = 'translation_ru.properties') -> list:
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
        with open('translation_ru.properties', 'r') as dict_file:
            for row in dict_file:
                if len(dict_value) < len(name_value_for_pars):
                    find_value = [x for x in name_value_for_pars if x in row]
                    if find_value:
                        # удаляем все символы до знака "="
                        dict_value[find_value[0]] = row.split('=', 1)[1].lstrip()
            dict_file.close()
    return dict_value


# decode значений роли
def decode(unicode_role: str) -> str:
    return unicode_role.encode().decode('unicode-escape')


# функция вытаскивает из архива если получится реализовать запуск на linux
def jar_unzip() -> str:
    archive = zipfile.ZipFile('integration-bundle.jar', 'r')
    archive.extractall('tmp_script')
    archive.close()
    return 'tmp_script//i18n//translation_ru.properties'


# подключение к серверу по SSH
def ssh_connect(server_ip, login, password):
    port = 22
    transport = paramiko.Transport((server_ip, port))
    transport.connect(username=login, password=password)
    sftp = paramiko.SFTPClient.from_transport(transport)
    print('ssh authentication is completed')
    remote_file_with_extensions = '/opt/ankey/ankey/extensions/gazprom-invest-integration-bundle-0.0-SNAPSHOT.jar'
    remote_file_with_ui = '/opt/ankey/ankey/ui/default/ng/public/i18n/translation_ru.properties'
    remote_file_with_sequential_approval_request = '/opt/ankey/ankey/conf/SequentialApprovalRequest.json'

    local_file_with_extensions = 'integration-bundle.jar'
    local_file_with_ui = 'translation_ru.properties'
    local_file_with_sequential_approval_request = 'SequentialApprovalRequest.json'

    sftp.get(remote_file_with_extensions, local_file_with_extensions)
    sftp.get(remote_file_with_ui, local_file_with_ui)
    sftp.get(remote_file_with_sequential_approval_request, local_file_with_sequential_approval_request)

    sftp.put(local_file_with_sequential_approval_request, remote_file_with_sequential_approval_request)
    sftp.put(local_file_with_extensions, remote_file_with_extensions)
    sftp.put(local_file_with_ui, remote_file_with_ui)
    sftp.close()
    transport.close()


# удаление временных файлов\папок
def remove_tmp_dir():
    path = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'tmp_script')
    file1 = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'integration-bundle.jar')
    file2 = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'SequentialApprovalRequest.json')
    file3 = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'translation_ru.properties')
    shutil.rmtree(path)
    os.remove(file1)
    os.remove(file2)
    os.remove(file3)


# определение имени интеграционного бандла
def find_name_bundle():
    pass


if __name__ == '__main__':
    ssh_connect(server_ip=input('IP address server Ankey IDM: '), login=input('login: '), password=getpass.getpass())
    main_func()
    remove_tmp_dir()

