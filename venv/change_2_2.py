import getpass
import json
import os
import shutil
import zipfile

import paramiko
from openpyxl import Workbook
from openpyxl.styles import Font


# работа с файлом excel
def excel_file(upload_json_data):
    wb = Workbook()
    ws = wb.active
    separator = '\u2192'  # красивый разделить для excel между ролями
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
def json_file():
    upload_data = []
    upload_json_data = {}
    with open('tmp//SequentialApprovalRequest.json', 'r', encoding='utf-8') as input_file:
        sequential_approval_request = json.load(input_file)  # cериализация json файла
    name_ir_list = list(sequential_approval_request.keys())
    url_translation_ru = jar_unzip()
    dictionary = create_dict(url_translation_ru)
    for name_ir in name_ir_list:
        excel_list = []
        if 'managerStage' in sequential_approval_request[name_ir].keys() and \
                sequential_approval_request[name_ir]['managerStage']['isEnabled'] == True:
            excel_list.append('Линейный руководитель')
        for stages in sequential_approval_request[name_ir]['stages']:
            stage = stages[0]
            if '$' in stage:
                field_name = sequential_approval_request[name_ir]['roleVariables'][stage]['fieldName']
                managed_object = \
                    sequential_approval_request[name_ir]['roleVariables'][stage]['managedObject'].split('/')[1]
                excel_list.append(decode(dictionary[managed_object + '.' + field_name]))
            else:
                excel_list.append(stage)
        upload_data.append(excel_list)
        upload_json_data[name_ir] = excel_list
    excel_file(upload_json_data)


def decode(unicode_role: str) -> str:
    return unicode_role.encode().decode('unicode-escape')


# создание словаря с unicode переменных из файлов транслитерации
def create_dict(url_translation_ru: str = 'translation_ru.properties') -> list:
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
def ssh_connect(server_ip, login, password):
    port = 22
    transport = paramiko.Transport((server_ip, port))
    transport.connect(username=login, password=password)
    sftp = paramiko.SFTPClient.from_transport(transport)
    print('ssh authentication is completed')
    remote_file_with_extensions = '/opt/ankey/ankey/extensions/gazprom-invest-integration-bundle-0.0-SNAPSHOT.jar'
    remote_file_with_ui = '/opt/ankey/ankey/ui/default/ng/public/i18n/translation_ru.properties'
    remote_file_with_sequential_approval_request = '/opt/ankey/ankey/conf/SequentialApprovalRequest.json'

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
    os.mkdir('tmp')


# удаление временных файлов\папок
def remove_tmp_dir():
    path = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'tmp')
    shutil.rmtree(path)


# определение имени интеграционного бандла
def find_name_bundle():
    pass


if __name__ == '__main__':
    create_tmp_dir()
    ssh_connect(server_ip=input('IP address server Ankey IDM: '), login=input('login: '), password=getpass.getpass())
    json_file()
    remove_tmp_dir()
