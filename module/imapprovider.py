import imaplib
import email
import re
import datetime

import base64
import quopri
import module.exceptionlite as exlite
from socket import error as socket_error
from collections import namedtuple

'''' Контейнер настроек подключения'''
ImapConfig = namedtuple('ImapConfig',
                        'name, \
                         host, \
                         port, \
                         login, \
                         password'
                        )
'''Контейнер адрессов'''
Contact = namedtuple('Contact',
                     'name, \
                      address'
                    )
'''Контейнер  сообщения'''
MessageData = namedtuple('MessageData', 
                         'address_from, \
                          address_to, \
                          address_cc, \
                          subject, \
                          body_plain, \
                          body_html, \
                          date_time, \
                          attachment'
                        )



'''  Основной класс для работы с почтой (IMAP)'''
class IMAPProvaider(object):
  '''
  Конструктор
    in: config(ImapConfig) Конфигурация
    throw: exlite.ExceptionLite
  '''
  def __init__(self, config):
    try:
      # Подключаемся к хосту
      self.mail = imaplib.IMAP4_SSL(config.host, config.port)
      # Авторизация
      self.mail.login(config.login, config.password)
      # Обновление списка папок IMAP
      self.__syncFolders__()
    # Перехват ошибок "транспорта"(socket)
    except socket_error as s:
      raise exlite.ExceptionLite('(Socket) ' + self.__bytesToStr__(s.args[1]))
    # Перехват ошибок IMAP
    except imaplib.IMAP4.error as i:
      raise exlite.ExceptionLite('(IMAP)' + self.__bytesToStr__(i.args[0]))
  '''
  Деструктор
  '''
  def __del__(self):
    try:
      if 'AUTH' == self.mail.state:
        self.mail.logout()
      self.mail.close()
    except:
      pass
  
  '''
  Синхронизация списка папок IMAP
  throw: ExceptionLite
  '''
  def __syncFolders__(self):
    try:
      # Получение данных с сервера
      (status, data) = self.mail.list()
      # Проверка статуса
      if 'OK' != status:
        raise exlite.ExceptionLite('status: ' + status)
      # Псевдонимы каталогов для коректного поиска и отображения
      dic_folders = {'inbox' : 'Входящие',
                     'important' : 'Важные',
                     'drafts' : 'Черновики',
                     'draftbox' : 'Черновики',
                     'sent' : 'Отправленные',
                     'sentbox' : 'Отправленные',
                     'flagged' : 'Помеченные',
                     'trash' : 'Корзина',
                     'junk' : 'Спам',
                     'spam' : 'Спам'
                    }
      # Список для очистки от неиспользуемых данных
      replace_flag_list = ('\HasNoChildren',
                           '\HasChildren',
                           'Noselect',
                           'All',
                           'Unmarked',
                           '\\',
                           '"',
                           ' '
                          )
      # Регулярное выражение для разбивки данных с сервера
      re_pattern = re.compile(r'\((?P<flags>.*?)\) "(?P<delimiter>.*)" (?P<name>.*)')
      self.folders = {}
      # Проход в цикле по всем строкам полученных с сервера
      for line in data:
        # Разбиваем согласно регулярки
        match = re_pattern.match(str(line, 'utf-8'))
        # Групируем
        (flags, delimiter, name) = match.groups()
        # Убираем кавычки и неиспользуемые элементы
        for replace_str in replace_flag_list:
          flags = flags.replace(replace_str, '')
          name = name.replace(replace_str, '')
        # Поиск имени пакпки в словаре
        if name.lower() in dic_folders:
          flags = dic_folders[name.lower()]
        else:
          # Если в словаре нет пробуем декодировать из Base64
          re_patern = re.compile('&(\S+)-')
          rlist = re.findall(re_patern, name)
          if 0 < len(rlist):
            flags = self.__decodeB64toStr__(rlist[0], 'utf-16BE')
        if '' != flags:
          self.folders[flags] = name
    # Перехват ошибок "транспорта"
    except socket_error as s:
      raise exlite.ExceptionLite('(Socket) ' + self.__bytesToStr__(s.args[1]))
    # Перехват ошибок IMAP
    except imaplib.IMAP4.error as i:
      raise exlite.ExceptionLite('(IMAP)' + self.__bytesToStr__(i.args[0]))
  
  '''
  Получение списка папок IMAP
  return: list - псевдо имена каталогов
  '''
  def getFoldersList(self):
    return sorted(list(self.folders.keys()))
  
  '''
  Установка папки IMAP для работы
  in: folder - псевдо имя папки из списка getFoldersList()
  throw: exlite.ExceptionLite
  '''
  def setFolder(self, folder):
    try:
      (status, data) = self.mail.select(self.folders[folder])
      if 'OK' != status:
        raise exlite.ExceptionLite('status: ' + status)
    # Перехват ошибок ввода
    except KeyError as ke:
      raise exlite.ExceptionLite('Folder not found: <%s>' % folder)
    # Перехват ошибок "транспорта"(socket)
    except socket_error as s:
      raise exlite.ExceptionLite('(Socket) ' + self.__bytesToStr__(s.args[1]))
    # Перехват ошибок IMAP
    except imaplib.IMAP4.error as i:
      raise exlite.ExceptionLite('(IMAP)' + self.__bytesToStr__(i.args[0]))
  
  '''
  Получение UID сообщений по выбранной папки
  in: new_only(bool) Только новые
  in: date_filter (datetime) [необязательный] возвращает информацию с определенной даты,
    без параметра возвращает информацию за все время
  return: list()
  throw: ExceptionLite
  '''
  def getUidList(self, new_only, date_filter = None):
    try:
      if None != date_filter:
        filter = 'SENTSINCE %s' % date_filter.strftime("%d-%b-%Y")
      else:
        filter = None
      if new_only:
        type_uid = 'UNSEEN'
      else:
        type_uid = 'ALL'
      # Получение данных в зависимости от фильтров
      (status, data) = self.mail.uid('search', None, filter, type_uid)
      # Проверка успешности получение данных
      if 'OK' != status:
        raise exlite.ExceptionLite('Status: %s' % (status))
      # Конвертируем данные
      uid_test = str(data[0], 'utf8')
      if '' == uid_test:
        return list()
      else:
        return str(data[0], 'utf8').split(' ')
    # Перехват ошибок "транспорта"(socket)
    except socket_error as s:
      raise exlite.ExceptionLite('(Socket) ' + self.__bytesToStr__(s.args[1]))
    # Перехват ошибок IMAP
    except imaplib.IMAP4.error as i:
      raise exlite.ExceptionLite('(IMAP)' + self.__bytesToStr__(i.args[0]))
  
  '''
  Получение  сообщения по UID полученного из getUidList()
  in: UID Уникальный номер сообщения
  return: (MessageData) Данные письма
  throw: ExceptionLite
  '''
  def getMessageData(self, uid):
    try:
      (status, data) = self.mail.uid('fetch', str(uid), '(RFC822)')
      # Проверка успешности получение данных
      if "OK" != status:
        raise exlite.ExceptionLite("Massage not found")
      # Конвертируем данные из bytes в Message
      msg = email.message_from_bytes(data[0][1])
      # Получаем "отправителя"
      address_from = self.__decodeAddress__(msg["From"])
      # Получаем список "получателей"
      address_to = self.__decodeAddress__(msg["To"])
      # Получаем список "в копии"
      address_cc= self.__decodeAddress__(msg["Cc"])
      # Дата и время отправки сообщения
      date_time = self.__decodeTimestamp__(email.utils.parsedate_tz(msg['date']))
      # Тема письма
      subject = self.__decodeImapStr__(msg["subject"])
      # Тело письма и вложение(я)
      body_plain = ''
      body_html = ''
      attachment = {}
      if msg.is_multipart():
        for part in msg.walk():
          content_type = part.get_content_type()
          content_disposition = str(part.get('Content-Disposition'))
          content_charset = part.get_content_charset()
          if None == content_charset:
            content_charset = 'utf-8'
          if 'attachment' not in content_disposition:
            if 'text/plain' == content_type:
              body_plain = part.get_payload(decode=True).decode(content_charset)
            elif 'text/html'  == content_type:
              body_html = part.get_payload(decode=True).decode(content_charset)
          else:
            filename = self.__decodeImapStr__(part.get_filename())
            attachment[filename] = part.get_payload(decode=True)
      else:
        content_type = msg.get_content_type()
        content_disposition = str(msg.get("Content-Disposition"))
        content_charset = msg.get_content_charset()
        if None == content_charset:
          content_charset = 'utf-8'
        if 'attachment' not in content_disposition:
          if 'text/plain'  == content_type:
            body_plain = msg.get_payload(decode=True).decode(content_charset)
          elif 'text/html' == content_type:
            body_html = msg.get_payload(decode=True).decode(content_charset)
        else:
          filename = self.__decodeImapStr__(msg.get_filename())
          attachment[filename] = msg.get_payload(decode=True)
      # Возврат сообщения
      return MessageData(address_from,
                         address_to,
                         address_cc,
                         subject,
                         body_plain,
                         body_html,
                         date_time,
                         attachment
                        )
    # Перехват ошибок "транспорта"(socket)
    except socket_error as s:
      raise exlite.ExceptionLite('(Socket) ' + self.__bytesToStr__(s.args[1]))
    # Перехват ошибок IMAP
    except imaplib.IMAP4.error as i:
      raise exlite.ExceptionLite('(IMAP)' + self.__bytesToStr__(i.args[0]))
    
  '''
  Удаление сообщение по UID
  in: (uid) - уникальный номер сообщения
  throw: ExceptionLite
  '''
  def deleteMessage(self, uid):
    try:
      folder_trash = self.folders['Корзина']
      self.mail.uid('STORE',uid, '+FLAGS', '(\\Deleted)')
      self.mail.expunge()
    # Перехват ошибок ввода
    except KeyError as ke:
      raise exlite.ExceptionLite('Folder not found: <%s>' % ke.args[0])
    # Перехват ошибок "транспорта"(socket)
    except socket_error as s:
      raise exlite.ExceptionLite('(Socket) ' + self.__bytesToStr__(s.args[1]))
    # Перехват ошибок IMAP
    except imaplib.IMAP4.error as i:
      raise exlite.ExceptionLite('(IMAP)' + self.__bytesToStr__(i.args[0]))

  '''
  Декодирование Адресов
  К сожалению не смог воспользоваться decode_header в связи с несовместимостью
  форматов(кодировки) mail.ru, yandex.ru, rambler.ru
  in: (str) raw данные из масива msg
  return (list(Contact))
  '''
  def __decodeAddress__(self, raw_data):
    result = []
    if None == raw_data:
      return result
    # Очистка строки от паразитирующих символов
    raw_str = re.sub(r"<|>|\r|\t|\n|", '', raw_data)
    raw_str = re.sub(r"\,", ' ', raw_str)
    # Разделение на группы
    adrr_split_list = self.__decodeImapStr__(raw_str).split(' ')
    # Удадение пустых элементов после split
    adrr_split_list = list(filter(None, adrr_split_list))
    temp_name = ''
    # Разделение по правилу (отправитель - адрес)
    for test_str in adrr_split_list:
      re_test = re.search(r'\S+@\S+', test_str)
      if None == re_test:
        temp_name += test_str + ' '
        continue
      else:
        result.append(Contact(temp_name.rstrip(), test_str))
        temp_name = ''
    return result
  
  '''Декодирование Даты и времени письма'''
  def __decodeTimestamp__(self, timestamp):
    # Дата может быть пустой(например в черновиках)
    if None == timestamp:
      return datetime.datetime(1970, 1, 1, 0, 0, 0)
    (year, month, day, hour, minute, second) = timestamp[:6]
    return datetime.datetime(year, month, day, hour, minute, second)
  
  '''Декодирование IMAP строки формата =?Кодировка?Тип?Строка?='''
  def __decodeImapStr__(self, raw_data):
    if None == raw_data:
      return str()
    # Данные могут быть представленны несколькими кодированными строками
    list_name = raw_data.split(' ')
    # Тестовая регулярка для разбиения строки (Кодировка, Тип, Данные)
    re_pattern = re.compile(r'^=\?(\S+)\?(.)\?(\S+)\?=')
    result =''
    for raw_str in list_name:
      match = re_pattern.match(raw_str)
      if None == match:
        result += ' ' +raw_str
        continue
      # Групируем
      (char_code, data_type, data) = match.groups()
      # Декодируем из Base64
      if 'B' == data_type:
        result += self.__decodeB64toStr__(data, char_code)
      # Декодирование из Quoted Printable
      elif 'Q' == data_type:
        result += self.__decodeQuotedPrintabletoStr__(data, char_code)
    return result

  '''Декодирование из Base64 строки в строку заданной кодировки'''
  def __decodeB64toStr__(self, data, char_code):
    b64_str = data
    # Нормализация
    b64_str = b64_str.replace(',', '/')
    b64_str += "=" * ((4 - len(b64_str) % 4) % 4)
    # Декодирование
    b64_data_bytes = str.encode(b64_str)
    data_byte = base64.b64decode(b64_data_bytes)
    data_str = data_byte.decode(char_code)
    return data_str
  
  '''Декодирование из Quoted Printable строки в строку заданной кодировки'''
  def __decodeQuotedPrintabletoStr__(self, data, char_code):
    decoded_string=quopri.decodestring(data)
    data_str = decoded_string.decode(char_code)
    return data_str
  
  '''Декодирование из bytes строки в строку'''
  def __bytesToStr__(self, bytes_str):
    if isinstance(bytes_str, bytes):
      return bytes_str.decode()
    else:
      return bytes_str