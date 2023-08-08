import requests, socks, socket, logging, math, random
import json, re, time, openpyxl, os.path
from bs4 import BeautifulSoup as BS
from fake_useragent import UserAgent
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


socks.set_default_proxy(socks.SOCKS5, "localhost", 9150)
socket.socket = socks.socksocket
logging.basicConfig(level=logging.ERROR, filename="parsing_error.log",
                    filemode="a", format="%(asctime)s %(levelname)s %(message)s", encoding='cp1251')


# HOST = "https://sbis.ru/"
URL = "https://sbis.ru/contragents/"
INPUT_FILE_NAME = 'psfp5.xlsx'
OUT_FILE_NAME = 'output.xlsx'
LOG_INN = 'INN_with_error-logging.xlsx'
# encoding='cp1251'


def generator_from_file():
    if not os.path.isfile(INPUT_FILE_NAME):
        print("\033[91m\033[1m{}\033[0m".format("Входной файл отсутствует"))
        return
    book = openpyxl.load_workbook(filename=INPUT_FILE_NAME)
    sheet = book.active
    for row in sheet.values:
        # match = re.fullmatch(r"\d{12}", str(row[0]))
        # if not match:
        #     continue
        # yield match[0]
        match = re.fullmatch(r"\d{10}", str(row[0]))
        if not match:
            continue
        yield match[0]
    book.close()
    return


def get_ip():
    r = requests.get('https://api.ipify.org?format=json').json()
    return r['ip']


def generator_mask(previous='37.204.193.182'):
    ua = UserAgent().chrome
    while True:
        ip = get_ip()
        if previous != ip:
            previous = ip
            ua = UserAgent().chrome
            print(f'IP: {ip}\nuser-agent: {ua}')
        yield ua


def get_response_sbis(inn_local, m = iter(generator_mask())):
    '''
    TODO: Необходимо устранить проблему: HTTPSConnectionPool(host='sbis.ru', port=443): Max retries exceeded with url
    https://stackoverflow.com/questions/23013220/max-retries-exceeded-with-url-in-requests
    :param inn_local: ИНН: str
    :return: response
    '''
    while True:
        try:
            session = requests.Session()
            retry = Retry(connect=7, backoff_factor=math.exp(1)+random.uniform(-0.01, 0.01)) # Настройка частоты запросов
            adapter = HTTPAdapter(max_retries=retry)
            session.mount('https://', adapter)
            response_local = session.get(URL + inn_local + "/",
                                          headers={"accept": "*/*", 'user-agent': next(m)}) # timeout=10
            if response_local.status_code == 200:
                return response_local
            # else:
            #     print("\033[96m{}\033[0m".format("sleep"*5))
            #     time.sleep(30.)
        except (socks.SOCKS5Error, socks.GeneralProxyError, requests.exceptions.ConnectionError) as ex:
            logging.error(f"INN: {inn_local} - {ex}")
            print("\033[91m\033[1m{}\033[0m".format(f"RESPONSE {ex}")) # : This error does`t go to the log file
            waiting_time = 60. # math.exp(6-retry)
            print("\033[96m{}\033[0m".format(f"sleep {waiting_time} seconds"))
            time.sleep(waiting_time)
            # writing_error_inn(inn_local)
            # return None


def get_content_sbis(response, inn_local=''):
    '''
    Информация по регулярным выражениям:
    https://pythonist.ru/regulyarnye-vyrazheniya-v-python/
    https://habr.com/ru/post/349860/
    https://habr.com/ru/company/ods/blog/346632/
    https://youtube.com/playlist?list=PLA0M1Bcd0w8w8gtWzf9YkfAxFCgDb09pA
    :param response:
    :param inn_local: ИНН: str
    :return:
    '''
    soup = BS(response.text, 'lxml').find('div', class_='wasabyJSDeps').findAll('script', type="text/javascript")
    final_dict = dict()
    for line in soup:
        if line.get_text() != '':
            word = line.get_text()[26:-2].replace('\\\\', "").replace('\"', "")
            break
    # if bool(re.findall(r'(ликвид|ЛИКВИД|Ликвид)', word)):
    #     print(word)
    #     print("\033[91m\033[1m{}\033[0m".format(f"INN: {inn_local} - The company has been liquidated / Компания находится в ликвидации"))
    #     return {'company': 'Ликвидирована', 'inn': inn_local}

    try:
        match = re.search(r'_type:record,d:\[([\sа-яА-Я0-9ёЁ«»+.-]+),([\sа-яА-Я0-9-+]+)', word)
        final_dict.update({'company': match[1] + match[2], 'inn': inn_local})

        match = re.findall(r'phone:\{count:\d,items:\[([^\]]+)', word)
        final_dict.update({'phone': match[0] if bool(match) else ''})

        match = re.findall(r'email:\{count:\d{1,2},items:\[([^\]]+)', word)
        final_dict.update({'email': match[0] if bool(match) else ''})

        match = re.findall(r'site:\{count:\d,items:\[([^\]]+)', word)
        final_dict.update({'site': match[0] if bool(match) else ''})

        match = re.findall(r',([\sа-яА-ЯёЁ-]+),(?:Генеральный Директор|Директор|Президент|Управляющая компания),', word)
        final_dict.update({'fio': match[0] if bool(match) else ''})

        match = re.findall(r'([\sа-яА-ЯёЁ-]+)', word)
        final_dict.update({'okved': match[match.index('Количество филиалов') + 1],
                           'address': ' '.join(list(re.findall(r',(\d{6}), ([\s\wа-яА-ЯёЁ0-9№\)\(.,/-]+),\{email', word)[0])),
                           'date_of_creation': re.findall(r'short:(?:Действует|В состоянии реорганизации, действует)\s+с\s+([\d.]+)[^,]', word.replace('\"', ""))[0]
                           })

        match = re.findall(r'short:(?:Действует|В состоянии реорганизации, действует)\s+с\s+.+{([\d.,:]+)\},\{', word)
        final_dict.update({'revenue': match[0] if bool(match) else ''})

        match = re.findall(r'\},null,(null|[\d.-]+),(null|[\d.-]+),(?:[\d-]+),(?:[\d.-]+),(?:[\d.-]+),(?:[\d.-]+),null,\{_type:recordset', word.replace(' ', ''))
        final_dict['r_sales'] = f'ROS: {match[0][0]}%' if bool(match) else ''
        final_dict['r_capital'] = f'ROE: {match[0][1]}%' if bool(match) else ''


        # print(json.dumps(final_dict, indent=4, ensure_ascii=False))
    except (IndexError, TypeError, ValueError) as ex:
        # print(json.dumps(final_dict, indent=4, ensure_ascii=False))
        writing_error_inn(inn_local)
        print("\033[91m\033[1m{}\033[0m".format(f"INN: {inn_local} - FILTER_FUNCTION: {ex}"))
        logging.error(f"INN: {inn_local} - FILTER_FUNCTION: {ex}")
        # print(json.dumps(final_dict, indent=4, ensure_ascii=False))
    return final_dict


def create_out_file() -> None:
    if os.path.isfile(OUT_FILE_NAME):
        return
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "data"
    sheet['A1'].value = 'Название (компания)'
    sheet['B1'].value = 'ИНН'
    sheet['C1'].value = 'Рабочий телефон (компания)'
    sheet['D1'].value = 'Рабочий email (компания)'
    sheet['E1'].value = 'Web'
    sheet['F1'].value = 'Адрес (компания)'
    sheet['G1'].value = 'Дата создания (компания)'
    sheet['H1'].value = 'Примечание'
    book.save(OUT_FILE_NAME)
    book.close()


def writing_to_out_file(d_info=None) -> None:
    book = openpyxl.load_workbook(filename=OUT_FILE_NAME)
    sheet = book.active
    list_info = []
    try:
        if d_info:
            list_info.append(d_info['company'])
            list_info.append(d_info['inn'])
            list_info.append(d_info['phone'])
            list_info.append(d_info['email'])
            list_info.append(d_info['site'])
            list_info.append(d_info['address'])
            list_info.append(d_info['date_of_creation'])

            rv = re.findall(r'\d+', d_info['revenue'])
            d_info['revenue'] = f'Выручка: {round((int(rv[-1]) + int(rv[-3]))/2//10**6)} млн.\n' if len(rv) >= 4 else 'Выручка: 0 млн.\n'
            d_info['revenue'] = '\n'.join([d_info['revenue'],
                                           "ОКВЭД: " + d_info['okved'],
                                           "Гендир: " + d_info['fio'],
                                           d_info['r_sales'],
                                           d_info['r_capital']
                                           ])
            list_info.append(d_info['revenue'])
            sheet.append(list_info)
    except Exception as ex:
        # print(json.dumps(final_dict, indent=4, ensure_ascii=False))
        writing_error_inn(d_info['inn'])
        print("\033[91m\033[1m{}\033[0m".format(f"INN: {d_info['inn']} - WRITING_TO_OUT_FILE: {ex}"))
        logging.error(f"INN: {d_info['inn']} - FILTER_FUNCTION: {ex}")
    finally:
        book.save(OUT_FILE_NAME)
        book.close()


def writing_for_analytics(d_info=None) -> None:
    # match = re.findall(r'short:(?:Действует|В состоянии реорганизации, действует)\s+с\s+.+{([\d.,:]+)\},\{([\d.,:]+)\}.+{([\d.,:]+)\}', word)
    # print(match)
    try:
        book = openpyxl.load_workbook(filename=OUT_FILE_NAME)
        sheet = book.active
        list_info = []
        if list_info['company'] == 'Ликвидирована':
            list_info.append(d_info['company'])
            list_info.append(d_info['inn'])
            d_info=None
        if d_info:
            list_info.append(d_info['company'])
            list_info.append(d_info['inn'])
            list_info.append(d_info['date_of_creation'])
            list_info.append(d_info['r_sales'])
            list_info.append(d_info['r_capital'])
            list_info.append(d_info['okved'])
        sheet.append(list_info)
        book.save(OUT_FILE_NAME)
        book.close()
    except Exception as ex:
        # print(json.dumps(final_dict, indent=4, ensure_ascii=False))
        writing_error_inn(d_info['inn'])
        print("\033[91m\033[1m{}\033[0m".format(f"INN: {d_info['inn']} - FILTER_FUNCTION: {ex}"))
        logging.error(f"INN: {d_info['inn']} - FILTER_FUNCTION: {ex}")
        # print(json.dumps(final_dict, indent=4, ensure_ascii=False))
    finally:
        book.save(OUT_FILE_NAME)
        book.close()


def writing_error_inn(inn) -> None:
    if os.path.isfile(LOG_INN):
        book = openpyxl.load_workbook(filename=LOG_INN)
        sheet = book.active
    else:
        book = openpyxl.Workbook()
        sheet = book.active
        sheet.title = "logging"
        sheet['A1'].value = 'ИНН'
    sheet.append([inn, ])
    book.save(LOG_INN)
    book.close()


def main():
    create_out_file()
    m = iter(generator_mask())
    for inn in generator_from_file():
        print(f'Работаем с ИНН: {inn}')
        if inn == '':
            return
        if len(inn) == 12:
            writing_to_out_file({'name': 'ИП', 'inn': inn})
            continue
        response = get_response_sbis(inn, m)
        content = get_content_sbis(response, inn)
        if content is not None:
            writing_to_out_file(content)
            # writing_for_analytics(content)
            print("Контент записан в файл")

        print("{s}{w}{s}".format(w="<end-of-iteration>", s='—'*25))
    print("\033[91m\033[1m{s}{w}{s}\033[0m".format(w="<THE_END>", s='—'*30))


if __name__ == "__main__":
    main()
    # error_list = ['7716590120', '9710069317']
    # m = iter(generator_mask())
    # for inn in error_list:
    #     response = get_response_sbis(inn, m)
    #     content = get_content_sbis(response, inn)
    #     print(content['r_sales'], content['r_capital'])


