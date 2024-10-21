import json
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.keys import Keys
from copy import deepcopy
from datetime import datetime, timedelta
import pandas
import asyncio
import re
from messages import *
import pyperclip

global admin_com, ADMINS, browser_task, rassl_week_days, rassl_time, working, user_com, new_result_db


try:
    result_db = pandas.read_excel('result.xlsx')
    for column in ['Номер', 'Имя ребенка', 'Возраст ребенка', 'День', 'Время', 'Итог']:
        result_db[column] = result_db[column].astype(str)
except:
    df = pandas.DataFrame(columns=['Номер', 'Имя ребенка', 'Возраст ребенка', 'День', 'Время', 'Итог'])
    excel_writer = pandas.ExcelWriter('result.xlsx', engine='xlsxwriter')
    df.to_excel(excel_writer, index=False, sheet_name='result', freeze_panes=(1, 0))
    excel_writer._save()

if not os.path.isdir("userdata"):
     os.mkdir("userdata")

browser_task = 'chill'
# ADMIN = '+79890804510'
ADMIN = '+79999967109'
ADMINS = ['+79890804510', '+79999967109']
# ADMINS = ['+7 989 080-45-10']
admin_com = {}

working = False
with open('first_message_text', 'r', encoding='utf-8') as fm_file:
    message_text = fm_file.read()

with open('rassl_info', 'r', encoding='utf-8') as fr_file:
    se_time, f_days = fr_file.read().split('\n')
    se_time = list(map(int, se_time.split(' ')))
    rassl_time = [datetime.now().replace(hour=se_time[0], minute=se_time[1], second=0, microsecond=0),
                  datetime.now().replace(hour=se_time[2], minute=se_time[3], second=0, microsecond=0)]
    if f_days:
        rassl_week_days = list(map(int, f_days.split(' ')))
    else:
        rassl_week_days = []


class Group:

    def __init__(self, name):
        self.name = name
        self.age = []
        self.load_rasp()
        self.days = {}

    def is_have_lesson(self, day, time):
        if day in self.days.keys():
            if time in self.days[day]:
                return True
        return False

    def get_days_msg(self):
        msg = f'Актуальное расписание для группы {self.name} ({self.age[0]} - {self.age[1]} лет):\n'

        allclear = True
        for day, times in self.days.items():
            if day != 'Возраст':
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
                group_ages = {'первые шаги': [4, 5], 'простые механизмы': [5, 6], 'программирование Scratch': [6, 9], 'начальная робототехника': [6, 7],
                              'город роботов': [8, 9], 'спайк': [9, 10], 'техномир': [10, 11], 'лаборатория роботов': [11, 14]}
                need = True
                for group, age in group_ages.items():
                    write_zero_info[group] = {}
                    write_zero_info[group]['Возраст'] = age
                    for day in ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье']:
                        write_zero_info[group][day] = []
        if need:
            with open('groups.json', 'w', encoding='utf-8') as file:
                json.dump(write_zero_info, file, indent=4, ensure_ascii=False)

        with open('groups.json', 'r', encoding='utf-8') as file:
            all_rasp = json.load(file)
            self.days = all_rasp[self.name]
            self.age = all_rasp[self.name]['Возраст']


groups = [Group(name) for name in ['первые шаги', 'простые механизмы', 'программирование Scratch', 'начальная робототехника', 'город роботов', 'спайк', 'техномир', 'лаборатория роботов']]

web_wait = lambda element, search_type, path, wait_time=10: WebDriverWait(element, wait_time).until(
    expected_conditions.presence_of_element_located((search_type, path)))


def convert_to_e164(phone_number):
    # Убираем все лишние символы, оставляем только цифры
    clean_number = re.sub(r'\D', '', str(phone_number))

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
def get_phone_numbers_from_excel():
    files = os.listdir(os.path.abspath(__file__)[:-8])
    table = ''
    for file in files:
        if '.xlsx' in file and file != 'result.xlsx':
            table = file
            break

    if table:
        df = pandas.read_excel(table)
        phone_numbers = []

        for column_name in df.columns.tolist():
            mbnumbers = df[column_name].astype(str).dropna().tolist()
            if mbnumbers and convert_to_e164(mbnumbers[0]):
                phone_numbers = [convert_to_e164(i.split(',')[0] if ',' in i else i) for i in mbnumbers]

        if phone_numbers:
            return [phone_numbers, 'ok']
        else:
            return [[], 'not found']
    else:
        return [[], 'no table']


# Функция для отправки сообщений
async def send_message(to_number, text):
    global browser, browser_task
    """
    :param to_number: Номер телефона получателя в формате E.164
    :param text: Текст сообщения
    """
    while browser_task != 'chill':
        await asyncio.sleep(2)

    browser_task = 'send'
    url = f"https://web.whatsapp.com/send?phone={to_number}&text="
    browser.get(url)

    try:
        text_area = web_wait(browser, By.XPATH, '//div[@aria-placeholder="Введите сообщение"]', 20)
        pyperclip.copy(text)
        text_area.send_keys(Keys.CONTROL, 'v')
        text_area.send_keys(Keys.ENTER)
        browser.find_element(By.XPATH, '//div[@id="main"]//div[@title="Меню"]').click()
        browser.find_element(By.XPATH, '//div[@id="app"]/div/span[5]//li[3]').click()
        browser_task = 'chill'
        return 'ok'
    except:
        try:
            browser.find_element(By.XPATH, '//div[@role="dialog"]//button').click()
            browser_task = 'chill'
            return 'Номер не действителен'
        except:
            browser_task = 'chill'
            return 'Неизвестная ошибка'


def save_db():
    with pandas.ExcelWriter('result.xlsx', engine='xlsxwriter') as writer:

        result_db.to_excel(writer, index=False, sheet_name='result')

        # Получаем объект workbook и worksheet
        workbook = writer.book
        worksheet = writer.sheets['result']

        # Применяем текстовый формат к столбцу 'Номер'
        text_format = workbook.add_format({'num_format': '@'})  # '@' означает текст
        worksheet.set_column('A:A', None, text_format)


# Основная функция для отправки сообщений по таймеру в указанные дни
async def send_messages_in_interval():
    global browser, browser_task, result_db, working
    # Получаем список номеров телефонов
    while True:
        phone_numbers, msg = get_phone_numbers_from_excel()
        if msg == 'ok':
            break
        if msg == 'no table':
            print('Таблица с номерами не найдена!')
            await asyncio.sleep(10)
        if msg == 'not found':
            print('В строках таблицы не найдены номера, убедитесь что вы загрузили нужную таблицу')
            await asyncio.sleep(10)
        if msg == 'error':
            print('Неизвестная ошибка, обратитесь к программисту')
            await asyncio.sleep(10)

    print(f'Найдено номеров разных лидов: {len(phone_numbers)}')

    try:
        wasnumbers = ['+' + i for i in result_db['Номер'].tolist()]
    except:
        wasnumbers = []

    # Цикл отправки сообщений
    for number in phone_numbers:
        while not working:
            print("Рассылка приостановлена")
            await asyncio.sleep(10)

        if number in wasnumbers:
            print(f'На номер {number} уже отправлено')
            continue
        # Получаем текущее время и день недели
        now = datetime.now()
        weekday = now.weekday()

        if weekday in rassl_week_days and rassl_time[0] <= now <= rassl_time[1]:
            # Отправляем сообщение
            res = await send_message(number, message_text)
            if res == 'ok':
                result_db.loc[len(result_db.index)] = [number, '-', '-', '-', '-', 'Ожидание ввода имени и возраста ребенка']  # adding a row
                print(f"Отправлено сообщение на номер {number}")
            else:
                result_db.loc[len(result_db.index)] = [number, '-', '-', '-', '-', res]
                print(f"Ошибка отправки на номер {number}")

            result_db = result_db.astype(str)
            save_db()

            # Ожидаем перед отправкой следующего сообщения
            await asyncio.sleep(60 * 5)
        else:
            if weekday not in rassl_week_days:
                print("На сегодняшний день нет рассылки")
                next_start = now + timedelta(days=(7 - weekday))
            else:
                print("На сегодня рассылка окончена")
                next_start = now + timedelta(days=1)

            next_start = next_start.replace(hour=rassl_time[0].hour, minute=rassl_time[1].minute, second=0, microsecond=0)
            sleep_time = (next_start - now).total_seconds()
            await asyncio.sleep(sleep_time)


# Функция для получения сообщений из чатов
async def get_messages():
    global browser, browser_task
    """
    Получает новые сообщения из чатов WhatsApp.

    :return: Словарь с номерами телефонов и списком новых сообщений
    """

    while browser_task != 'chill':
        await asyncio.sleep(2)

    browser_task = 'check'
    messages_dict = {}

    # Найти все чаты на боковой панели
    chats = browser.find_elements(By.XPATH, '//div[@aria-label="Список чатов"]/div')
    chats = [[int(chat.value_of_css_property('transform').split(' ')[-1][:-1]), chat] for chat in chats]
    chats = sorted(chats, key=lambda point: (point[0]))
    opened = False
    for loc, chat in chats:
        info = chat.text.split('\n')
        if len(info) == 4:
            try:
                need_messages = int(info[3])
            except Exception:
                continue
            chat.click()
            web_wait(browser, By.XPATH, '//div[@title="Сведения профиля"]').click()
            number = ''
            while number == '':
                number = browser.find_element(By.XPATH, f'//section//span[contains(text(), "+")]').text
                await asyncio.sleep(0.02)
            number = convert_to_e164(number)
            browser.find_element(By.XPATH, f'//div[@aria-label="Закрыть"]').click()
            messages_dict[number] = []
            opened = True
            xpath = '//div[@id="main"]/div[3]/div/div[2]/div[@role="application"]'
            messages = web_wait(browser, By.XPATH, xpath)
            otdel = messages.find_elements(By.XPATH, '//div[@role="row"]')
            for i in range(len(otdel), len(otdel) - need_messages, -1):
                messages_dict[number].append(otdel[i - 1].text[:-6])
            messages_dict[number].reverse()

    if opened:
        browser.find_element(By.XPATH, '//div[@id="main"]//div[@title="Меню"]').click()
        while True:
            try:
                web_wait(browser, By.XPATH, '//div[@id="app"]/div/span[5]//li[3]').click()
                break
            except:
                await asyncio.sleep(1)

    browser_task = 'chill'
    return messages_dict


def start_browser():
    global browser
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
    with open('rassl_info', 'r', encoding='utf-8') as r_file:
        ttime, ddays = r_file.read().split('\n')
    for day in list(map(int, ddays.split(' '))):
        msg += f"{days_list[day]}, "
    msg = msg[:-2] + '\n'
    ttime = ttime.split(' ')
    msg += f'Время рассылки: {ttime[0]}:{ttime[1]}-{ttime[2]}:{ttime[3]}\n' \
           f'Состояние рассылки: {"Запущена" if working else "Остановлена"}'
    return msg


async def work_with_getted_messages():
    global browser_task, admin_com, message_text, rassl_week_days, rassl_time, working, user_com
    while True:
        await asyncio.sleep(5)

        msg_dict = await get_messages()

        if msg_dict:
            for number, msgs in msg_dict.items():
                print(number, msgs)

                if number in ADMINS:
                    if number.startswith('+7'):
                        number = convert_to_e164(number)
                    if number not in admin_com.keys():
                        msgs[0] = msgs[0].lower()
                        if msgs[0] == 'получить текст рассылки':
                            await send_message(number, message_text)

                        elif msgs[0] == 'изменить текст рассылки':
                            await send_message(number, f'Отправьте новый текст')
                            admin_com[number] = 'new-text'

                        elif msgs[0].startswith('получить расписание'):
                            if msgs[0] == 'получить расписание':
                                group = ''
                                msg = 'Расписание для всех групп:'
                            else:
                                group = msgs[0][len('получить расписание '):]
                                msg = ''

                            for need_group in groups:
                                if group in need_group.number:
                                    msg += need_group.get_days_msg() + '\n\n'

                            await send_message(number, msg)

                        elif msgs[0].startswith('изменить расписание'):
                            if msgs[0] == 'изменить расписание':
                                await send_message(number, WRONG_GROUP)
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
                                admin_com[number] = f'new-raspis_{group}_{day}'
                                await send_message(number, msg)

                        elif msgs[0].startswith('получить информацию о рассылке'):
                            await send_message(number, get_rassl_info())

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

                                    await send_message(number, f'Время рассылки установлено на {newtime}')
                                except:
                                    await send_message(number, WRONG_RASSL_TIME_MSG)
                            else:
                                await send_message(number, WRONG_RASSL_TIME_MSG)

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
                            await send_message(number, get_rassl_info())

                        elif msgs[0] == 'запустить рассылку':
                            working = True
                            await send_message(number, get_rassl_info())

                        elif msgs[0] == 'остановить рассылку':
                            working = False
                            await send_message(number, get_rassl_info())

                        else:
                            await send_message(number, ADMINS_MSG)
                        continue

                    elif admin_com[number] == 'new-text':
                        message_text = msgs[0]
                        with open('first_message_text', 'w', encoding='utf-8') as file:
                            file.write(message_text)
                        await send_message(number, f'Текст рассылки изменён: \n{message_text}')
                        del admin_com[number]

                    elif admin_com[number].startswith('new-raspis'):
                        group, day = admin_com[number][len('new-raspis_'):].split('_')
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
                                if ngroop.number == group:
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
                                if ngroop.number == group:
                                    ngroop.load_rasp()
                                    msg += ngroop.get_days_msg()

                        await send_message(number, msg)
                        del admin_com[number]
                else:
                    all_numbers = result_db['Номер'].tolist()
                    user_com = ''
                    need_i = -1
                    for i in range(len(all_numbers)):
                        if all_numbers[i] == number:
                            user_com = result_db['Итог'].tolist()[i]
                            need_i = i

                    if not user_com:
                        print('\n'.join(msgs))
                    elif user_com == 'Ожидание ввода имени и возраста ребенка':
                        try:
                            c_name, age = msgs[0].split(' ')
                            age = int(age)
                            msg = ''
                            for group in groups:
                                if group.age[0] <= age <= group.age[1]:
                                    msg += group.get_days_msg() + '\n\n'
                            if not msg:
                                await send_message(number, 'Спасибо за ваш ответ! К сожалению для детей такого возраста у нас нет занятий.')
                            else:
                                msg += 'Пожалуйста, введите день недели и время, в которое вы бы хотели посетить занятие. Например:\n' \
                                       'Понедельник 17:30'
                                result_db.at[need_i, 'Имя ребенка'] = c_name
                                result_db.at[need_i, 'Возраст ребенка'] = age
                                result_db.at[need_i, 'Итог'] = 'Ожидание выбора дня и времени'
                                await send_message(number, msg)

                        except:
                            await send_message(number, 'Пожалуйста, проверьте что вы ввели данные в правильном формате! Пример правильного формата:\n'
                                                       'Иван 9')
                    elif user_com == 'Ожидание выбора дня и времени':
                        try:
                            day, time = msgs[0].split(' ')
                            time = time.replace('.', ':')
                            finded = False
                            for group in groups:
                                if group.age[0] <= result_db['Возраст ребенка'].tolist()[need_i] <= group.age[1]:
                                    finded = group.is_have_lesson(day, time)
                                    if finded:
                                        break

                            if finded:
                                result_db.at[need_i, 'День'] = day
                                result_db.at[need_i, 'Время'] = time
                                result_db.at[need_i, 'Итог'] = 'Записаны'

                                msg = f'Спасибо за ответ! Вы записаны в {day} на {time}, будем вас ждать!\n' \
                                      f'Если появятся какие-либо вопросы, можете написать их сюда, я обязательно сообщу администратору и с вами свяжутся!'
                                await send_message(number, msg)
                                await send_message(ADMIN, f'Пользователь {number} записался в {day} на {time}')
                            else:
                                await send_message(number, 'Пожалуйста, проверьте что вы ввели данные в правильном формате и указали правильное время! Пример правильного формата:\n'
                                                           'Понедельник 17:30')

                        except:
                            await send_message(number, 'Пожалуйста, проверьте что вы ввели данные в правильном формате! Пример правильного формата:\n'
                                                       'Понедельник 17:30')

                    elif user_com == 'Записаны':
                        await send_message(ADMIN, f'Пользователь {number} написал сообщениу:\n'
                                                  f'{". ".join(msgs)}')

            save_db()


async def undermain():
    global browser
    browser = start_browser()
    await asyncio.gather(work_with_getted_messages(), send_messages_in_interval())


# Пример использования
if __name__ == "__main__":
    asyncio.run(undermain())
