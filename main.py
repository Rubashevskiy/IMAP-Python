import datetime
import os
import pathlib
import module.exceptionlite as exlite
from datetime import timedelta

from module.imapprovider import IMAPProvaider as ipr
from module.imapprovider import ImapConfig as ipr_cfg


cfg_gmail_com = ipr_cfg('gmail.com','imap.gmail.com', 993, 'login', 'password')
cfg_mail_ru = ipr_cfg('mail.ru','imap.mail.ru', 993, 'login', 'password')
cfg_yandex_ru = ipr_cfg('yandex.ru','imap.yandex.ru', 993, 'login', 'password')
cfg_rambler_ru = ipr_cfg('rambler.ru','imap.rambler.ru', 993, 'login', 'password')

def saveMailAttachment(path, filename, data):
  try:
    if not os.path.isdir(path):
      pathlib.Path(path).mkdir(parents=True, exist_ok=True)
    filepath = os.path.join(path, filename)
    open(filepath, "wb").write(data)
  except PermissionError as e:
    raise ipr.IMAPException("Permission Error" + str(e))

def ContactListToStr(contact_list):
  result =''
  count = len(contact_list)
  curent = 0
  for contact in contact_list:
    if 0 != len(contact.name):
      result += '%s <%s>' % (contact.name, contact.address)
    else:
      result += '<%s>' % (contact.address)
    curent += 1
    if curent < count:
      result += ', '
  return result

def getMail(cfg):
  try:
    s_sep = '#'*60
    print(s_sep)
    print('Запуск конфигурации: %s' % cfg.name)
    print('Сервер:              %s' % cfg.host)
    print('Ящик:                %s' % cfg.login)
    date_time_start = datetime.datetime.now()
    print('Время:               %s' % date_time_start)
    print(s_sep)
    mail =  ipr(cfg)
    m_sep = '-'*60
    for folder in mail.getFoldersList():
      print('Каталог: %s' % folder)
      mail.setFolder(folder)
      for uid in mail.getUidList(False):
        msg = mail.getMessageData(uid)
        print('  ОТ: %s' % (ContactListToStr(msg.address_from)))
        print('  Кому: %s' % (ContactListToStr(msg.address_to)))
        if 0 < len(msg.address_cc):
          print('  Копия: %s' % (ContactListToStr(msg.address_cc)))
        print('  Тема: %s' % (msg.subject))
        print('  Дата и время: %s' % (msg.date_time))
        if 0 < len(list(msg.attachment.keys())):
          print('  Файлы: %s' % (list(msg.attachment.keys())))
        print(m_sep)
    print(s_sep)
    print('Завершенно успешно')
    date_time_end = datetime.datetime.now()
    print('Время:               %s' % date_time_end)
    print('Затраченное времени: %s' % str(date_time_end - date_time_start))
    print(s_sep)
    print('\n\r\n\r')
  except exlite.ExceptionLite as e:
    print('ERROR: ', e.what)
    e.PrintTraceback()
    exit(1)

def main():
  imap_config_list = [cfg_gmail_com, cfg_mail_ru, cfg_yandex_ru, cfg_rambler_ru]
  for cfg in imap_config_list:
    getMail(cfg)
  exit(0)

if __name__ == "__main__":
  main()