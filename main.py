import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.keys import Keys
from copy import deepcopy
import time
from datetime import datetime, timedelta
import pandas as pd
import asyncio
import re
from messages import *

global admin_com, ADMINS, browser_task, rassl_week_days, rassl_time, working, user_com

browser_task = 'chill'
ADMIN = '+79890804510'
#ADMINS = ['+7 989 080-45-10', '+7 999 996-71-09']
ADMINS = ['+7 989 080-45-10']
admin_com = {}

with open('usercom.json', 'r', encoding='utf-8') as file:
    user_com = json.load(file)

working = True
with open('first_message_text', 'r', encoding='utf-8') as file:
    message_text = file.read()

with open('rassl_info', 'r', encoding='utf-8') as file:
    se_time, days = file.read().split('\n')
    se_time = list(map(int, se_time.split(' ')))
    rassl_time = [datetime.now().replace(hour=se_time[0], minute=se_time[1], second=0, microsecond=0),
                  datetime.now().replace(hour=se_time[2], minute=se_time[3], second=0, microsecond=0)]
    if days:
        rassl_week_days = list(map(int, days.split(' ')))
    else:
        rassl_week_days = []


class Group:

    def __init__(self, name):
        group_ages = {'пе': [4, 5], 'пр': [5, 6], 'на': [6, 7], 'го': [7, 9], 'те': [10, 14], 'сп': [10, 14]}
        self.name = name
        self.age = group_ages[name[:2]]
        self.load_rasp()

    def get_days_msg(self):
        msg = f'Актуальное расписание для группы {self.name} ({self.age[0]} - {self.age[1]} лет):\n'

        allclear = True
        for day, times in self.days.items():
            if times:
                allclear = False
                msg += f'{day}: '
                for ttime in times:
                    msg += f'{ttime}, '
                msg = msg[:-2] + '\n'

        if allclear:
            msg += 'Нет занятий'

        return msg

    def load_rasp(self):
        write_zero_info = {}
        need = False
        with open('groups.json', 'r', encoding='utf-8') as file:
            if file.read() == '':
                need = True
                for group in ['первые шаги', 'простые механизмы', 'начальная робототехника', 'город роботов', 'техномир', 'спайк']:
                    write_zero_info[group] = {}
                    for day in ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье']:
                        write_zero_info[group][day] = []
        if need:
            with open('groups.json', 'w', encoding='utf-8') as file:
                json.dump(write_zero_info, file, indent=4, ensure_ascii=False)

        with open('groups.json', 'r', encoding='utf-8') as file:
            all_rasp = json.load(file)
            self.days = all_rasp[self.name]


groups = [Group('первые шаги'), Group('простые механизмы'), Group('начальная робототехника'), Group('город роботов'), Group('техномир'), Group('спайк')]

web_wait = lambda element, search_type, path, wait_time=10: WebDriverWait(element, wait_time).until(
    expected_conditions.presence_of_element_located((search_type, path)))


def convert_to_e164(phone_number):
    """
    Преобразует телефонный номер в формат E.164.

    :param phone_number: Телефонный номер (строка)
    :param country_code: Код страны (например, "7" для России)
    :return: Номер в формате E.164 или None, если номер некорректен
    """
    # Убираем все лишние символы, оставляем только цифры
    clean_number = re.sub(r'\D', '', phone_number)

    # Если номер уже начинается с кода страны (например, +7 или 7 для России)
    if clean_number.startswith('7'):
        return f"+{clean_number}"

    elif clean_number.startswith('8'):
        return f"+7{clean_number[1:]}"

    # Если номер не содержит кода страны, добавляем его
    elif len(clean_number) == 10:  # Проверка на локальный формат (без кода страны)
        return f"+7{clean_number}"

    # Если номер не соответствует никакому формату, возвращаем None
    else:
        return None


# Функция для получения номеров телефонов из Excel
def get_phone_numbers_from_excel(file_path, column_name):
    """
    Функция для получения номеров телефонов из столбца Excel.

    :param file_path: Путь к Excel-файлу
    :param column_name: Название столбца с номерами телефонов
    :return: Список номеров телефонов
    """
    #return ['+7(928)455-95-55']
    return ['+79999967109']

    df = pd.read_excel(file_path)
    phone_numbers = df[column_name].dropna().tolist()  # Убираем пустые значения
    return phone_numbers


# Функция для отправки сообщений
def send_message(to_number, text):
    global browser
    """
    Отправляет сообщение пользователю через Wazzap.

    :param to_number: Номер телефона получателя в формате E.164
    :param message_text: Текст сообщения
    """
    url = f"https://web.whatsapp.com/send?phone={to_number}&text="
    browser.get(url)

    btn = '//button[@aria-label="Отправить"]'
    text_area = web_wait(browser, By.XPATH, '//div[@aria-placeholder="Введите сообщение"]', 20)
    for line in text.split('\n'):
        text_area.send_keys(Keys.SHIFT, Keys.ENTER)
        text_area.send_keys(line)
    text_area.send_keys(Keys.ENTER)
    browser.find_element(By.XPATH, '//div[@id="main"]//div[@title="Меню"]').click()
    browser.find_element(By.XPATH, '//div[@id="app"]/div/span[5]//li[3]').click()


# Основная функция для отправки сообщений по таймеру в указанные дни
async def send_messages_in_interval(file_path, column_name):
    global browser
    """
    Отправляет сообщения по списку номеров раз в минуту с 8:30 до 22:00
    с понедельника по пятницу.

    :param file_path: Путь к Excel-файлу с номерами телефонов
    :param column_name: Название столбца с номерами телефонов
    :param message_text: Текст сообщения
    """
    global browser_task
    # Получаем список номеров телефонов
    phone_numbers = get_phone_numbers_from_excel(file_path, column_name)

    # Цикл отправки сообщений
    for number in phone_numbers:
        # Получаем текущее время и день недели
        now = datetime.now()
        weekday = now.weekday()

        if weekday in rassl_week_days and rassl_time[0] <= now <= rassl_time[1] and working:
            # Отправляем сообщение
            while browser_task != 'chill':
                await asyncio.sleep(2)
            browser_task = 'send'
            send_message(number, message_text)
            user_com[number] = 'get-name-and-age'
            save_userscom()
            browser_task = 'chill'
            print(f"Отправлено сообщение на номер {number}")

            # Ожидаем одну минуту перед отправкой следующего сообщения
            await asyncio.sleep(60 * 3)
        else:
            if weekday not in rassl_week_days:
                print("На сегодняшний день нет рассылки")
                next_start = now + timedelta(days=(7 - weekday))
            elif not working:
                print("Рассылка приостановлена")
                next_start = now + timedelta(days=1)
            else:
                print("На сегодня рассылка окончена")
                next_start = now + timedelta(days=1)

            next_start = next_start.replace(hour=rassl_time[0].hour, minute=rassl_time[1].minute, second=0, microsecond=0)
            sleep_time = (next_start - now).total_seconds()
            await asyncio.sleep(sleep_time)


# Функция для получения сообщений из чатов
def get_messages():
    global browser
    """
    Получает новые сообщения из чатов WhatsApp.

    :param driver: Экземпляр драйвера Selenium
    :return: Словарь с номерами телефонов и списком новых сообщений
    """
    messages_dict = {}

    # Найти все чаты на боковой панели
    chats = browser.find_elements(By.XPATH, '//div[@aria-label="Список чатов"]/div')
    opened = False
    for chat in chats:
        info = chat.text.split('\n')
        if len(info) == 4:
            opened = True
            messages_dict[info[0]] = []
            need_messages = int(info[3])
            chat.click()
            xpath = '//div[@id="main"]/div[3]/div/div[2]/div[@role="application"]'
            messages = web_wait(browser, By.XPATH, xpath)
            otdel = messages.find_elements(By.XPATH, '//div[@role="row"]')
            for i in range(len(otdel), len(otdel) - need_messages, -1):
                messages_dict[info[0]].append(otdel[i - 1].text[:-6])
            messages_dict[info[0]].reverse()

    if opened:
        browser.find_element(By.XPATH, '//div[@id="main"]//div[@title="Меню"]').click()
        browser.find_element(By.XPATH, '//div[@id="app"]/div/span[5]//li[3]').click()

    return messages_dict


def start_browser():
    options = Options()
    options.binary_location = 'C:\Program Files\Google\Chrome\Application\chrome.exe'
    options.add_argument('--allow-profiles-outside-user-dir')
    options.add_argument('--enable-profile-shortcut-manager')
    options.add_argument(r'user-data-dir=C:\Users\Кирилл\PycharmProjects\Robotex\userdata')  # УКАЖИТЕ ПУТЬ ГДЕ ЛЕЖИТ ВАШ ФАЙЛ. Советую создать отдельную папку.
    options.add_argument('--profile-directory=Profile 1')
    options.add_argument('--profiling-flush=n')
    options.add_argument('--enable-aggressive-domstorage-flushing')
    options.add_argument('--log-level=3')
    options.add_argument('--enable-chrome-browser-cloud-management')
    headless_option = deepcopy(options)
    headless_option.add_argument('--headless')
    url = f"https://web.whatsapp.com"
    browser = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    browser.get(url)
    try:
        web_wait(browser, By.XPATH, '//canvas[@aria-label="Scan this QR code to link a device!"]')
        print('Отсканируй QR')
        web_wait(browser, By.ID, 'pane-side', 60)
    finally:
        return browser


def get_rassl_info():
    days_list = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс']
    msg = 'Дни, в которые осуществляется рассылка: '
    with open('rassl_info', 'r', encoding='utf-8') as file:
        ttime, days = file.read().split('\n')
    for day in list(map(int, days.split(' '))):
        msg += f"{days_list[day]}, "
    msg = msg[:-2] + '\n'
    ttime = ttime.split(' ')
    msg += f'Время рассылки: {ttime[0]}:{ttime[1]}-{ttime[2]}:{ttime[3]}\n' \
           f'Состояние рассылки: {"Запущена" if working else "Остановлена"}'
    return msg


def save_userscom():
    with open('usercom.json', 'w', encoding='utf-8') as file:
        json.dump(user_com, file, indent=4, ensure_ascii=False)


def db_write(name, c_name, age):
    pass

def db_write_dt(name, day, time):
    pass

async def work_with_getted_messages():
    global browser_task, admin_com, message_text, rassl_week_days, rassl_time, working
    await asyncio.sleep(5)
    while True:
        msg_dict = {}
        if browser_task == 'chill':
            browser_task = 'check'
            msg_dict = get_messages()
            browser_task = 'chill'

        if msg_dict:
            for name, msgs in msg_dict.items():
                if name in ADMINS:
                    if name.startswith('+7'):
                        name = convert_to_e164(name)
                    if name not in admin_com.keys():
                        msgs[0] = msgs[0].lower()
                        if msgs[0] == 'получить текст рассылки':
                            send_message(name, message_text)

                        elif msgs[0] == 'изменить текст рассылки':
                            send_message(name, f'Отправьте новый текст')
                            admin_com[name] = 'new-text'

                        elif msgs[0].startswith('получить расписание'):
                            if msgs[0] == 'получить расписание':
                                group = ''
                                msg = 'Расписание для всех групп:'
                            else:
                                group = msgs[0][len('получить расписание '):]
                                msg = ''

                            for need_group in groups:
                                if group in need_group.name:
                                    msg += need_group.get_days_msg() + '\n\n'

                            send_message(name, msg)

                        elif msgs[0].startswith('изменить расписание'):
                            if msgs[0] == 'изменить расписание':
                                send_message(name, WRONG_GROUP)
                                continue
                            else:
                                group_day = (msgs[0][len('изменить расписание '):] + ' ').split(' ')
                                if 'спайк' in group_day or 'техномир' in group_day:
                                    group = group_day[0]
                                    day = group_day[1]
                                else:
                                    group = group_day[0] + ' ' + group_day[1]
                                    day = group_day[2]

                                if day == '':
                                    msg = f'Введите новое расписание для группы "{group}". Правильный формат:\n\n' \
                                          f'Понедельник 13:45 15:15 17:00\n' \
                                          f'Четверг 11:00\n\n' \
                                          f'(не прописанные дни будут считаться пустыми, вместо двоеточия можно использовать точку)'
                                else:
                                    msg = f'Введите новое расписание для группы "{group}" на {day}. Правильный формат:\n\n' \
                                          f'13:45 15:15 17:00\n\n' \
                                          f'(вместо двоеточия можно использовать точку)'
                                admin_com[name] = f'new-raspis_{group}_{day}'
                                send_message(name, msg)

                        elif msgs[0].startswith('получить информацию о рассылке'):
                            send_message(name, get_rassl_info())

                        elif msgs[0].startswith('изменить время рассылки'):
                            newtime = msgs[0].split(' ')[-1]
                            tint = []
                            if 8 < len(newtime) < 12:
                                try:
                                    for ttime in newtime.split('-'):
                                        tint.append(int(ttime[:-3]))
                                        tint.append(int(ttime[-2:]))
                                    with open('rassl_info', 'r', encoding='utf-8') as file:
                                        days = file.read().split('\n')[1]
                                    with open('rassl_info', 'w', encoding='utf-8') as file:
                                        file.write(' '.join(list(map(str, tint))) + '\n' + days)

                                    rassl_time = [datetime.now().replace(hour=tint[0], minute=tint[1], second=0, microsecond=0),
                                                  datetime.now().replace(hour=tint[2], minute=tint[3], second=0, microsecond=0)]

                                    send_message(name, f'Время рассылки установлено на {newtime}')
                                except:
                                    send_message(name, WRONG_RASSL_TIME_MSG)
                            else:
                                send_message(name, WRONG_RASSL_TIME_MSG)

                        elif msgs[0].startswith('изменить дни рассылки'):
                            days_list = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс']
                            days = msgs[0][len('изменить дни рассылки'):]
                            rassl_week_days = []
                            for check_day in range(7):
                                if days_list[check_day] in days:
                                    rassl_week_days.append(check_day)
                            with open('rassl_info', 'r', encoding='utf-8') as file:
                                time = file.read().split('\n')[0]
                            with open('rassl_info', 'w', encoding='utf-8') as file:
                                file.write(' '.join(list(map(str, rassl_week_days))) + '\n' + time)
                            send_message(name, get_rassl_info())

                        elif msgs[0] == 'запустить рассылку':
                            working = True
                            send_message(name, get_rassl_info())

                        elif msgs[0] == 'остановить рассылку':
                            working = False
                            send_message(name, get_rassl_info())

                        else:
                            send_message(name, ADMINS_MSG)
                        continue

                    elif admin_com[name] == 'new-text':
                        message_text = msgs[0]
                        with open('first_message_text', 'w', encoding='utf-8') as file:
                            file.write(message_text)
                        send_message(name, f'Текст рассылки изменён: \n{message_text}')
                        del admin_com[name]

                    elif admin_com[name].startswith('new-raspis'):
                        group, day = admin_com[name][len('new-raspis_'):].split('_')
                        msg = 'Расписание изменено!\n\n'
                        if day == '':
                            new_raspis = {}
                            for line in msgs[0].split('\n'):
                                need_info = line.replace('.', ':').split(' ')
                                new_raspis[need_info[0][0].upper() + need_info[0][1:]] = need_info[1:]
                            with open('groups.json', encoding='utf-8') as file:
                                all_rasp = json.load(file)
                            for day, ttimme in new_raspis.items():
                                all_rasp[group][day] = ttimme
                            with open('groups.json', 'w', encoding='utf-8') as file:
                                json.dump(all_rasp, file, indent=4, ensure_ascii=False)
                            for ngroop in groups:
                                if ngroop.name == group:
                                    ngroop.load_rasp()
                                    msg += ngroop.get_days_msg()

                        else:
                            times = msgs[0].replace('.', ':').split(' ')
                            with open('groups.json', encoding='utf-8') as file:
                                all_rasp = json.load(file)
                            all_rasp[group][day[0].upper() + day[1:]] = times
                            with open('groups.json', 'w', encoding='utf-8') as file:
                                json.dump(all_rasp, file, indent=4, ensure_ascii=False)
                            for ngroop in groups:
                                if ngroop.name == group:
                                    ngroop.load_rasp()
                                    msg += ngroop.get_days_msg()

                        send_message(name, msg)
                        del admin_com[name]
                else:
                    if name.startswith('+7'):
                        name = convert_to_e164(name)
                    if name not in user_com.keys():
                        text = ''.join(msgs)
                    elif user_com[name] == 'get-name-and-age':
                        try:
                            c_name, age = msgs[0].split(' ')
                            age = int(age)
                            msg = ''
                            for group in groups:
                                if group.age[0] <= age <= group.age[1]:
                                    msg += group.get_days_msg() + '\n\n'
                            if not msg:
                                send_message(name, 'Спасибо за ваш ответ! К сожалению для детей такого возраста у нас нет занятий.')
                            else:
                                msg += 'Пожалуйста, введите день недели и время, в которое вы бы хотели посетить занятие. Например:\n' \
                                       'Понедельник 11.00'
                                db_write(name, c_name, age)
                                send_message(name, msg)
                                user_com[name] = 'get-day-and-time'

                        except:
                            send_message(name, 'Пожалуйста, проверьте что вы ввели данные в правильном формате! Пример правильного формата:\n'
                                               'Иван 9')
                    elif user_com[name] == 'get-day-and-time':
                        try:
                            day, time = msgs[0].split(' ')
                            db_write_dt(name, day, time)
                        except:
                            pass

            save_userscom()


async def undermain():
    global browser
    file_path = "phone_numbers.xlsx"  # Путь к вашему файлу Excel
    column_name = "Phone"  # Название столбца с номерами
    browser = start_browser()
    await asyncio.gather(work_with_getted_messages(), send_messages_in_interval(file_path, column_name))




# Пример использования
if __name__ == "__main__":
    asyncio.run(undermain())
