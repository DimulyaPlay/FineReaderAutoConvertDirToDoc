import os
import sys
from subprocess import call
from time import sleep

config_name = 'config.cfg'

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)

config_path = os.path.join(application_path, config_name)
with open(config_path, encoding = 'utf-8', mode = 'r') as conf:
    params = conf.read().splitlines()
finecmd_path = params[0].split(' = ')[1]
input_folder = fr'{params[1].split(" = ")[1]}'
output_folder = fr'{params[2].split(" = ")[1]}'
delay = params[3].split(' = ')[1]

print('Скрипт для мониторинга директории и конвертации файлов через FineReader')
print('Dmity Sosnin, Krasnogamsky gs, github.com/dumulyaplay')
print()
print()
print('Указан путь к FineCmd: ', finecmd_path)

if os.path.isfile(finecmd_path):
    if os.path.basename(finecmd_path) == 'FineCmd.exe':
        print('Подтвержден')
    else:
        print('файл не является finecmd.exe')
        input()
        sys.exit(0)
else:
    print('файл не является finecmd.exe')
    input()
    sys.exit(0)

print('Указан путь к папке со сканами: ', input_folder)
if os.path.isdir(input_folder):
    print('Путь подтвержден')
else:
    print('Путь не является папкой')
    input()
    sys.exit(0)

print('Указан путь к папке для отправки распознанных файлов: ', output_folder)
if os.path.isdir(output_folder):
    print('Путь подтвержден')
else:
    print('Путь не является папкой')
    input()
    sys.exit(0)

print('Указана периодичность сканирования, секунд:', delay)
try:
    delay = int(delay)
except:
    print('Периодичность должна быть числом и исчисляется в секундах')
    input()
    sys.exit(0)

current_files = os.listdir(input_folder)
sleep(delay)
print('Сканирование директории запущено')
while True:
    new_files = os.listdir(input_folder)
    new_current_files = [i for i in new_files if i not in current_files and i != os.path.basename(output_folder)]
    sleep(2)
    if new_current_files:
        print('найдены новые файлы: ', new_current_files)
        for new_file in new_current_files:
            filepath_in = input_folder + r'\\' + new_file
            filepath_out = output_folder + r'\\' + new_file + '.doc'
            if os.path.isfile(filepath_in):
                print('конвертация ', new_file, ' началась')
                call(finecmd_path + ' ' + filepath_in + ' /lang russian /out ' + filepath_out + ' /quit')
                print('конвертация ', new_file, ' завершена')
                sleep(2)
            else:
                print('пропущена папка: ', new_file)
        current_files = new_files
        print('Ожидание новых файлов')
    sleep(delay)
