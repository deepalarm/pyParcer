from pathlib import Path
import os, pdb
import xml.etree.ElementTree as ET
import win32com.client
import pythoncom
import shutil
# Показываем текущую директорию и переходим в директорию Загрузки текущего пользователя

print("Текущая директория", os.getcwd())
#pdb.set_trace()
fold_dir = str(Path.home())
os.chdir(fold_dir)
os.chdir('Downloads')
new_fold_dir=os.getcwd()
print("Директория изменена на:", new_fold_dir)  # Смотрим файлы в каталоге
for dirs, folders, files in os.walk(new_fold_dir):
    for file in files:
        if file.endswith('.xml'):   # Ищем файлы с расширением xml
            name_count = ''
            curr_filename = (file)
            print (' ')
            print ('Обрабатывается файл: ',curr_filename) # Выводим имя найденного файла
            tree = ET.parse(curr_filename)  # Парсим файл и ищем нужный тэг
            root = tree.getroot()
            xml_attr = (root.find('.//{*}DocType')) 
            print ('Тип документа: ',xml_attr.text)
            doc_type = xml_attr.text  # Определяем тип документа
            if str(doc_type) == 'O_IP_ACT_GACCOUNT_MONEY':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_REJECT_PROLONGIP':
                xml_attr = root.find('.//{*}Correspondent')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}DbtrName')
            elif doc_type == 'O_IP_ACT_ZP':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_RES_REOPEN':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_SVOD_DBT':
                xml_attr = root.find('.//{*}Correspondent')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}DbtrName')
            elif doc_type == 'O_IP_ACT_BAN_EXIT':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_ENDBAN_GIBDD':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_ENDGACCOUNT_MONEY':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_ARR_COSTDOWN':
                xml_attr = root.find('.//{*}Correspondent')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}DbtrName')
            elif doc_type == 'O_IP_ACT_STOP_OTHER':
                xml_attr = root.find('.//{*}Correspondent')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}DbtrName')
            elif doc_type == 'O_IP_ACT_ARREST_ACCMONEY':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_CURRENCY_ROUB':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_REOPEN_CANCEL':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_REOPEN':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_END_CANCEL':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_RES_SATISFY_PETITION':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_RETURN':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_PENS':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_DBT_CANCEL':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_END_END':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_END_STOP':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_ENDBAN_REG':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_BAN_REG':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_ARREST':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_BAN_GIBDD':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_ENDARREST':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_SVOD_COL':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_FINDDOCS_END':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_GM_BUDGET':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_FIX_SPI':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_REJECT_CHNAGESP':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_ARR_APROVCOST':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_REJECT_CHNAGESPI':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_END_REOPEN':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_CHARGE_CANCEL':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_ENDARR':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_STORE_PROP':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_REJECT_DEPRT_ARREST':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_ORDER_OTHER':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_TO_OTHER_OSP':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_CANCEL_FIND':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_DELAY':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_EXECUTE':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_FIND':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_GETCURRENCY':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_CHANGE_SIDE':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_RD_FINDDOCS_RDRISE':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_ARR_DEFCOST':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_RES_ANY':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_ENDSALE':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_RES_TAKERIGHTBECRDR_CN':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_NOWORK':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_ZEK':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_ARR_SETSTOR':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_ARR_SETESTIM':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_CANCEL':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_REJECT_REQ_SPI_CANCEL':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_DECLINEFIND':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_RES_ARREST_SELFSALE':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_ARR_GETAUC':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_FIND_ACCOUNT':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_ARR_TODEB':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_ARR_GETSALE':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_REJECT_STOPIP':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_REJECT_ENDIP':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_LIVING_WAGE':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')
            elif doc_type == 'O_IP_ACT_COMPENS':
                xml_attr = root.find('.//{*}DbtrName')
                if str(xml_attr) == 'None':
                    xml_attr = root.find('.//{*}Correspondent')



            else:
                print ('Документ с типом: ', doc_type , 'отсутствует в списке обработки')
                os.system('pause')
                continue
            print (xml_attr.text)
            debitor_name = xml_attr.text # В переменную debitor_name запихиваем ФИО должника
            try:    # Начинаем преобразование файла для дальнейшего поиска pdf
                pdf_find = str(curr_filename) # Присваиваем имя XML файла
                print ('Для файла:', pdf_find, 'ищем файл', end =' ')
                to_pdf = (pdf_find[0:-4] + '.pdf') # В переменную to_pdf помещаем значение имени файла вида - piev_12345678.pdf
                print (to_pdf)
                pdf_exist = str(os.path.exists(to_pdf)) # Проверяем наличие PDF файла вида - piev_12345678.pdf (Должны переименовать в ФИО.pdf)
                if str(pdf_exist) == 'True': # Нашли такой файл.
                    i = 1 # Итерационный счетчик
                    print ('Файл - ', to_pdf, 'найден, начинаем переименование')

                    if doc_type == 'O_IP_ACT_DBT_CANCEL':
                        if str(os.path.exists(debitor_name + ' - отмена мер на доходы' +'.pdf')) == 'False': # Проверяем есть ли файл ФИО.pdf
                            os.rename(to_pdf, debitor_name + ' - отмена мер на доходы' +'.pdf') # Если такого файла нет, переименовываем 
                            os.remove(curr_filename)
                        else:
                            while str(os.path.exists(debitor_name + name_count + ' - отмена мер на доходы' +'.pdf')) == 'True':
                                name_count = ' '+'(' + str(i) + ')'
                                print (name_count)
                                i+=1
                            else:
                                os.rename(to_pdf, debitor_name + name_count + ' - отмена мер на доходы' +'.pdf') # Если такого файла нет, переименовываем 
                                os.remove(curr_filename)
                        continue

                    if doc_type == 'O_IP_ACT_LIVING_WAGE':
                        if str(os.path.exists(debitor_name + ' - о сохранении заработной платы и иных доходов' + '.pdf')) == 'False':  # Проверяем есть ли файл ФИО.pdf
                            os.rename(to_pdf, debitor_name + ' - о сохранении заработной платы и иных доходов' + '.pdf')  # Если такого файла нет, переименовываем
                            os.remove(curr_filename)
                        else:
                            while str(os.path.exists(debitor_name + name_count + ' - о сохранении заработной платы и иных доходов' + '.pdf')) == 'True':
                                name_count = ' ' + '(' + str(i) + ')'
                                print(name_count)
                                i += 1
                            else:
                                os.rename(to_pdf, debitor_name + name_count + ' - о сохранении заработной платы и иных доходов' + '.pdf')  # Если такого файла нет, переименовываем
                                os.remove(curr_filename)
                        continue

                    if str(os.path.exists(debitor_name + '.pdf')) == 'False':
                        os.rename(to_pdf, debitor_name + '.pdf') # Если такого файла нет, переименовываем 
                        os.remove(curr_filename)
                    else: # Если присутствует файл вида ФИО.pdf
                        while str(os.path.exists(debitor_name + name_count + '.pdf')) == 'True':
                            print ('Внутри цикла while файл deb_name несет значение =', debitor_name)
                            name_count = ' '+'(' + str(i) + ')'
                            print (name_count)
                            i+=1
                        else:
                                
                            os.rename(to_pdf, debitor_name + name_count + '.pdf')
                            os.remove(curr_filename)
                            
                else:
                    print ('Отсутствует PDF файл с именем: ', to_pdf)
                    continue
                
            except:
                print ('Все пошло по пизде.')
print ('Обработка файлов закончена')

# Инициализурием COM подключение
conn_str = "Srvr='SERVER EITH DATA BASE';Ref='_______ ';Usr='TYPE USER NAME HERE';Pwd='TYPE PASSWORD HERE';"
pythoncom.CoInitialize()
V83 = win32com.client.Dispatch("V83.COMConnector").Connect(conn_str)
current_dir = str(new_fold_dir)
dir_to_copy = ('C:\Output destination')

print(current_dir)
print(dir_to_copy)

os.system('pause')

for dirs, folders, files in os.walk(new_fold_dir):
    for file in files:
        if file.endswith('.pdf'):
            curr_filename = file
            ShortFilename = curr_filename[0:-4]
            print (ShortFilename)
            NotFirstFile = ShortFilename.find('(')
            Decline_Income = ShortFilename.find('- отмена')
            print(NotFirstFile)
            if NotFirstFile > 0:
                ShortFilename = ShortFilename[0:NotFirstFile - 1]
            elif Decline_Income > 0:
                ShortFilename = ShortFilename[0:Decline_Income - 1]

            print(ShortFilename)
            print(Decline_Income)

            Manager = V83.Exchange.GetResponsibles(ShortFilename)
            print (Manager)
            srs_dir = current_dir + '\\' + curr_filename
            dst_dir = dir_to_copy + '\\' + Manager
            shutil.move(curr_filename, dst_dir)
os.system('pause')


