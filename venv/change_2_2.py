import getpass
import json
import os
import shutil
import zipfile
import paramiko
import argparse
from openpyxl import Workbook
from openpyxl.styles import Font

CONFIG_FILE = 'conf.json'
IS_DICT = 'is_name.json'
UNIQUE_DICT = 'unique_dict.json'

DEFAULT_SEPARATOR = ' \u2192 '
SINGLE_STAGE_SEPARATOR = ' || '


# агрументы для запуска скрипта
def arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument("f", help='если хотите запустить с уникальным словарем необходим аргумент F \n '
                                  'Если хотите запустить с стандартным словарем проекта необходим аргумент D')
    args_start = parser.parse_args()
    return args_start


# работа с файлом excel
def excel_file(upload_json_data):
    wb = Workbook()
    ws = wb.active
    separator = DEFAULT_SEPARATOR  # красивый разделить для excel между ролями
    # заголовки столбцов + выделение жирным шрифтом
    ws['A1'] = 'Имя информационного ресурса'
    ws['B1'] = 'Этапы согласования'
    ws['A1'].font = Font(bold=True)
    ws['B1'].font = Font(bold=True)
    # размеры слобцов
    ws.column_dimensions['A'].width = 50
    ws.column_dimensions['B'].width = 120
    row_offset = 2
    for i, keys in enumerate(upload_json_data.keys()):
        ws['A{}'.format(i + row_offset)] = keys
        ws['B{}'.format(i + row_offset)] = separator.join(upload_json_data[keys])
    wb.save("Этапы согласования.xlsx")
    print("Файл Этапы согласования.xlsx успешно создан в текущей директории")


# работа с файлом универсальной цепочки
def json_file(unique_dictionary: dict = {}, is_unique_dict: bool = False):
    upload_json_data = {}
    with open('tmp//SequentialApprovalRequest.json', 'r', encoding='utf-8') as input_file:
        sequential_approval_request = json.load(input_file)  # cериализация json файла
    # исключаем '_id' переменную из json цепочки(если цепочка была залита черезе rest)
    name_ir_list = [name for name in sequential_approval_request.keys() if name != '_id']
    # Встроенный справочник транслитерации этапов
    url_translation_ru = jar_unzip()
    dictionary = create_dict(url_translation_ru)
    for name_ir in name_ir_list:
        excel_list = []
        if 'managerStage' in sequential_approval_request[name_ir].keys() and \
                sequential_approval_request[name_ir]['managerStage']['isEnabled'] is True:
            excel_list.append('Линейный руководитель')
        for stages in sequential_approval_request[name_ir]['stages']:
            current_stage = []
            for stage in stages:
                if '$' in stage:
                    if is_unique_dict is True and stage in unique_dictionary.keys():
                        current_stage.append(unique_dictionary[stage])
                    else:
                        field_name = sequential_approval_request[name_ir]['roleVariables'][stage]['fieldName']
                        managed_object = \
                            sequential_approval_request[name_ir]['roleVariables'][stage]['managedObject'].split('/')[1]
                        current_stage.append(decode(dictionary[managed_object + '.' + field_name]))
                else:
                    current_stage.append(stage)
            excel_list.append(SINGLE_STAGE_SEPARATOR.join(current_stage))
            current_stage.clear()
          # обработка словаря с именами ИС
        is_dictionary = read_json(IS_DICT)
        if name_ir in is_dictionary.keys():
            upload_json_data[is_dictionary[name_ir]] = excel_list
        else:
            upload_json_data[name_ir] = excel_list

    excel_file(upload_json_data)


def decode(unicode_role: str) -> str:
    return unicode_role.encode().decode('unicode-escape')


# создание словаря с unicode переменных из файлов транслитерации
def create_dict(url_translation_ru: str = 'translation_ru.properties') -> dict:
    dict_value = {}
    in_symbol = 'templates.managed.form.'
    out_symbol = 'placeholder'
    with open(url_translation_ru, 'r') as dict_file:
        for row in dict_file:
            # ищем вхождение поля
            if in_symbol in row and out_symbol not in row:
                # делим по знаку '=' и сохраняем только тип объекта и его транслитерацию
                dict_value[row.split('=')[0].replace(in_symbol, '')] = row.split('=')[1].strip()
        dict_file.close()
    with open('tmp//translation_ru.properties', 'r') as dict_file:
        for row in dict_file:
            # ищем вхождение поля
            if in_symbol in row and out_symbol not in row:
                # делим по знаку '=' и сохраняем только тип объекта и его транслитерацию
                dict_value[row.split('=')[0].replace(in_symbol, '')] = row.split('=')[1].strip()
        dict_file.close()
    return dict_value


# функция вытаскивает из архива если получится реализовать запуск на linux
def jar_unzip() -> str:
    archive = zipfile.ZipFile('tmp//integration-bundle.jar', 'r')
    archive.extractall('tmp//tmp_script')
    archive.close()
    return 'tmp//tmp_script//i18n//translation_ru.properties'


# подключение к серверу и отправка
def ssh_connect(server_ip, login, password, ankey_dir, intergation_bundle_name):
    port = 22
    transport = paramiko.Transport((server_ip, port))
    transport.connect(username=login, password=password)
    sftp = paramiko.SFTPClient.from_transport(transport)
    print('ssh authentication is completed')
    remote_file_with_extensions = ankey_dir + "ankey/extensions/" + intergation_bundle_name
    remote_file_with_ui = ankey_dir + 'ankey/ui/default/ng/public/i18n/translation_ru.properties'
    remote_file_with_sequential_approval_request = ankey_dir + 'ankey/conf/SequentialApprovalRequest.json'
    local_file_with_extensions = 'tmp\\integration-bundle.jar'
    local_file_with_ui = 'tmp\\translation_ru.properties'
    local_file_with_sequential_approval_request = 'tmp\\SequentialApprovalRequest.json'

    sftp.get(remote_file_with_extensions, local_file_with_extensions)
    sftp.get(remote_file_with_ui, local_file_with_ui)
    sftp.get(remote_file_with_sequential_approval_request, local_file_with_sequential_approval_request)

    sftp.put(local_file_with_sequential_approval_request, remote_file_with_sequential_approval_request)
    sftp.put(local_file_with_extensions, remote_file_with_extensions)
    sftp.put(local_file_with_ui, remote_file_with_ui)
    sftp.close()
    transport.close()


# создание временной директории для файлов полученных по sftp
def create_tmp_dir():
    if os.path.exists('tmp'):
        shutil.rmtree('tmp')
    return os.mkdir('tmp')


# удаление временных файлов\папок
def remove_tmp_dir():
    return shutil.rmtree('tmp')


# универсальный обработчик json
def read_json(json_f) -> dict:
    try:
        with open(json_f, 'r', encoding='utf-8') as input_file:
            dict_file = json.load(input_file)  # cериализация json файла
        return dict_file
    except FileNotFoundError:
        dict_file = {}
        return dict_file


if __name__ == '__main__':
    args = arguments()
    default_cred = read_json(CONFIG_FILE)
    create_tmp_dir()
    ssh_connect(server_ip=default_cred['server_ip'], login=default_cred['login'], password=getpass.getpass(),
                ankey_dir=default_cred['ankey_dir'], intergation_bundle_name=default_cred['intergation_bundle_name'])
    if args.f == 'F':
        is_unique_dict = True
        unique_dictionary = read_json(UNIQUE_DICT)
        json_file(unique_dictionary, is_unique_dict)
    if args.f == 'D':
        json_file()
    remove_tmp_dir()
