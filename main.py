import imaplib
import email
from email.header import decode_header
import os
import re
from urllib.parse import unquote
from tqdm import tqdm

def clean_filename(filename):
    """
    Очищает имя файла от недопустимых символов и удаляет лишние подчеркивания.
    """
    filename = unquote(filename)  # Декодирование URL-кодированных символов
    filename = filename.replace('\0', '')  # Удаление символов null
    filename = re.sub(r'[<>:"/\\|?*\r\n\t]', '_', filename)  # Замена недопустимых символов на подчеркивания
    filename = re.sub(r'__+', '_', filename)  # Удаляет повторяющиеся подчеркивания
    return filename.strip('_')  # Удаляет подчеркивания с начала и конца строки

def decode_mime_words(s):
    """
    Декодирует строку в соответствии с MIME-заголовками.
    """
    if s is None:
        return ""
    return ''.join(
        word.decode(encoding or 'utf-8', errors='ignore') if isinstance(word, bytes) else word
        for word, encoding in decode_header(s)
    )


def shorten_subject(subject, max_length=50):
    """
    Сокращает тему письма, если её длина превышает max_length символов,
    добавляя в конец три восклицательных знака (!!!).
    """
    if len(subject) > max_length:
        return subject[:max_length].rstrip() + "!!!"
    else:
        return subject


def process_account(email_user, email_pass):
    """
    Обрабатывает учетную запись электронной почты, извлекая и сохраняя сообщения.
    """
    try:
        mail = imaplib.IMAP4_SSL(imap_host)
        mail.login(email_user, email_pass)
        mail.select('inbox')

        result, data = mail.search(None, 'ALL')
        if result == 'OK':
            messages = data[0].split()
            print(f"Всего сообщений для {email_user}: {len(messages)}")

            # Инициализация счетчика для текущего пользователя
            message_counter = 1

            for num in tqdm(messages, desc=f"Сохранение сообщений для {email_user}"):
                result, data = mail.fetch(num, '(RFC822)')
                if result == 'OK':
                    raw_email = data[0][1]
                    email_message = email.message_from_bytes(raw_email)

                    subject = decode_mime_words(email_message['Subject'])

                    # Создаем папку для пользователя и темы сообщения с уникальным номером
                    user_folder = os.path.join("emails", email_user.split('@')[0])
                    subject_safe = clean_filename(subject).rstrip()  # Удаление пробелов справа
                    subject_safe = shorten_subject(subject_safe)  # Сокращаем тему письма если необходимо

                    # Добавляем порядковый номер к имени папки
                    folder_name = os.path.join(user_folder, f"{message_counter}_{subject_safe}")
                    os.makedirs(folder_name, exist_ok=True)

                    # Увеличиваем счетчик после создания каждой папки
                    message_counter += 1

                    body = ""
                    content_type = None

                    # Обработка содержимого сообщения
                    if email_message.is_multipart():
                        for part in email_message.walk():
                            content_type = part.get_content_type()
                            content_disposition = str(part.get("Content-Disposition"))

                            if content_type in ["text/plain", "text/html"] and "attachment" not in content_disposition:
                                payload = part.get_payload(decode=True)
                                try:
                                    body = payload.decode('utf-8')
                                except UnicodeDecodeError:
                                    body = payload.decode('iso-8859-1')
                                if content_type == "text/html":
                                    break
                    else:
                        content_type = email_message.get_content_type()
                        payload = email_message.get_payload(decode=True)
                        try:
                            body = payload.decode('utf-8')
                        except UnicodeDecodeError:
                            body = payload.decode('iso-8859-1')

                    from_ = decode_mime_words(email_message['From'])
                    to_ = decode_mime_words(email_message['To'])
                    file_extension = "html" if content_type == "text/html" else "txt"
                    # Сохранение сообщения
                    with open(os.path.join(folder_name, f"message.{file_extension}"), 'w', encoding='utf-8') as f:
                        if content_type == "text/html":
                            f.write(f"От: {from_}\n<br> Кому:{to_}\n<br>--------------------------------------------------\n<br>{body}")
                        else:
                            f.write(f"От: {from_}\n<br>Кому:{to_}\n<br>--------------------------------------------------\n<br>{body}")

                    # Обработка вложений
                    attachments = [part for part in email_message.walk() if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None]
                    for part in attachments:
                        file_name = decode_mime_words(part.get_filename())
                        if file_name:
                            file_name = clean_filename(file_name)
                            file_path = os.path.join(folder_name, 'files', file_name)
                            #file_path = shorten_path(file_path)  # Проверяем и при необходимости сокращаем путь
                            os.makedirs(os.path.dirname(file_path), exist_ok=True)
                            with open(file_path, 'wb') as f:
                                f.write(part.get_payload(decode=True))
            print(f"Обработка сообщений для {email_user} завершена.\n")
        else:
            print(f"Не удалось получить сообщения для {email_user}.\n")
    except imaplib.IMAP4.error:
        print(f"Не могу авторизоваться в учетной записи {email_user} возможно неверный Логин или Пароль!\n "
              f"Не забывайте что новый пароль приложений начнет действовать только через 2–3 часа!\n "
              f"Так же возможно что в разделе 'Все настройки > Почтовые программы' не указано 'Разрешить доступ к почтовому ящику с помощью почтовых клиентов'\n")

# Настройки
imap_host = 'imap.yandex.ru'

# Проверяем наличие файла с учетными данными
email_list_path = 'down_email.txt'
if os.path.exists(email_list_path) and os.path.getsize(email_list_path) > 0:
    with open(email_list_path, 'r') as file:
        for line in file:
            email_user, email_pass = line.strip().split(';')
            process_account(email_user, email_pass)
else:
    email_user = input("Введите ваш e-mail: ")
    email_pass = input("Введите ваш пароль: ")
    process_account(email_user, email_pass)