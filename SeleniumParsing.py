import json, re, time, os, openpyxl
from platform import system
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException

# TODO: Когда появляется несколько вариантов с руспрофайла алгоритм просто пропускает компанию, не заполняя данные
"""
pip install selenium --upgrade
Проверять обновления https://chromedriver.chromium.org/ !
"""
PATH = r'C:\Users\Grom\Desktop\Python\Parsing\chromedriver_win32\chromedriver.exe'
options = Options()
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--disable-notifications")
# options.add_argument('--headless') # без графического интерфейса.
options.add_experimental_option("prefs", {
    "profile.default_content_setting_values.media_stream_mic": 1,
    "profile.default_content_setting_values.media_stream_camera": 1,
    "profile.default_content_setting_values.geolocation": 0,
    "credentials_enable_service": False,
    "profile.password_manager_enabled": False,
    "profile.default_content_setting_values.notifications": 2,
})

EX = (NoSuchElementException, StaleElementReferenceException, )
URL_SEARCH = 'https://duckduckgo.com'
INPUT_FILE_NAME = '12.xlsx'
OUT_FILE_NAME = '111.xlsx'
STATISTIC = dict()


def check_point():
    browser = Chrome(r'C:\Users\Grom\Desktop\Python\Parsing\chromedriver_win32\chromedriver.exe') # service=Service(PATH), options=options
    browser.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        'source': '''
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Array;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Promise;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Symbol;
      '''
    })
    browser.get('https://nowsecure.nl') # Нормальный тестер
    # browser.get('https://intoli.com/blog/not-possible-to-block-chrome-headless/chrome-headless-test.html')


def generator_from_selenium():
    if not os.path.isfile(INPUT_FILE_NAME):
        print("\033[91m\033[1m{}\033[0m".format("Входной файл отсутствует"))
        return
    book = openpyxl.load_workbook(filename=INPUT_FILE_NAME)
    sheet = book.active
    for row in sheet.values:
        if bool(str(row[0])):
            yield str(row[0])
    book.close()
    return


def searcher():
    '''
    Ссылки для ознакомления с Selenium:
    https://temofeev.ru/info/articles/ultimativnaya-shpargalka-po-selenium-s-python-dlya-avtomatizatsii-testirovaniya/
    Для исправления ошибок:
    https://testengineer.ru/oshibki-v-selenium-gajd-po-exceptions/
    Для клавиш:
    https://www.selenium.dev/selenium/docs/api/py/webdriver/selenium.webdriver.common.keys.html?highlight=keys
    :return: список ссылок:list
    '''
    list_of_requests = iter(generator_from_selenium())
    browser = Chrome(options=options, service=Service(PATH))
    browser.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        'source': '''
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Array;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Promise;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Symbol;
      '''
    })
    browser.get(URL_SEARCH)
    req = next(list_of_requests)
    browser.find_element(By.XPATH, '//*[@id="search_form_input_homepage"]').send_keys(req + ' сайт')
    browser.find_element(By.XPATH, '//*[@id="search_button_homepage"]').click()
    while bool(req):
        result = []
        for i in range(5):
            res_s = WebDriverWait(browser, 10, 1000).until(
                EC.presence_of_element_located((By.XPATH, f'//*[@id="r1-{i}"]/div[1]/div/a')))
            result.append(res_s.get_attribute('href'))
        yield result
        req = next(list_of_requests)
        element_s = browser.find_element(By.XPATH, '//*[@id="search_form_input"]')
        element_s.click() #
        element_s.send_keys(Keys.CONTROL + "A" + Keys.DELETE)
        element_s.send_keys(req + ' сайт')
        browser.find_element(By.XPATH, '//*[@id="search_button"]').click()


def condition(link=''):
    tuple_patterns = ('sbis.ru', 'rusprofile.ru', 'list-org.com', 'companies.rbc.ru', 'audit-it.ru', 'spark-interfax.ru',
                  'vbankcenter.ru', 'checko.ru', 'find-org.com', 'synapsenet.ru', 'e-ecolog.ru', 'innproverka.ru',
                  'fek.ru', 'zachestnyibiznes.ru', 'cataloxy.ru', 'focus.kontur.ru')
    for pattern in tuple_patterns:
        if pattern in link:
            STATISTIC[pattern] = STATISTIC.get(pattern, 0) + 1
            return True
    return False


def duplicate_filter(links):
    pattern = r'http[s]?://([\w.-]+)/'
    res = [re.findall(pattern, el)[0] if bool(re.findall(pattern, el)) else el for el in links]
    return list(set(res))


def filter_links(links):
    i = 0
    while i < len(links):
        if condition(link=links[i]):
            links.pop(i)
        else:
            i += 1
    return links


def generator_from_rusprofile():
    data = ['9701114965', '7713440222', '9723051490', '7717729350', '7727193911', '7729759864',
            '7725302192', '7703608821', '7708769580', '5029181937', '7701720240']
    for inn in data:
        yield inn
    return


def get_info_from_rusprofile():
    list_inn = iter(generator_from_selenium())
    browser = Chrome(options=options, service=Service(PATH))
    browser.get('https://www.rusprofile.ru/')
    browser.find_element(By.XPATH, '//*[@id="menu-personal-trigger"]/span').click()
    browser.implicitly_wait(3)
    browser.find_element(By.XPATH, '//*[@id="v-root"]/div/div[1]/div[3]/div[2]/div/input').send_keys('ЛОГИН')
    browser.find_element(By.XPATH, '//*[@id="v-root"]/div/div[1]/div[3]/div[3]/div/input').send_keys('ПАРОЛЬ')
    browser.find_element(By.XPATH, '//*[@id="v-root"]/div/div[1]/div[3]/div[4]/button').click()
    try:
        browser.implicitly_wait(3)
        browser.find_element(By.XPATH, '//*[@id="mw-sa"]/div/div[6]').click()
        # notify = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, f'//*[@id="mw-sa"]/div/div[6]')))
        # notify.click()
    except NoSuchElementException:
        print('NoSuchElementException - Едем дальше')

    req = next(list_inn)
    time.sleep(3)
    browser.find_element(By.XPATH, '//*[@id="indexsearchform"]/div/input').send_keys(req)
    browser.find_element(By.XPATH, '//*[@id="indexsearchform"]/button').click()

    while bool(req):
        print(f'Работаем с ИНН: {req}')
        browser.implicitly_wait(3)
        str_page = browser.page_source.replace(r' ', '')
        telephones = re.findall(r'itemprop="telephone">([\d\s\)\(+-]+)</a>', str_page)
        emails = re.findall(r'itemprop="email">([\w@.-]+)</a>', str_page)
        print(telephones, emails, sep='\n')
        telephones = ["Телефоны:"] + telephones
        emails = ["email:"] + emails
        yield telephones + emails
        req = next(list_inn)
        search_form = browser.find_element(By.XPATH, '//*[@id="searchform"]/input[1]')
        search_form.clear()
        search_form.send_keys(req + Keys.ENTER)
        print("{s}{w}{s}".format(w="<end-of-iteration>", s='—'*25))
    print("\033[91m\033[1m{s}{w}{s}\033[0m".format(w="<THE_END>", s='—'*30))


def create_file() -> None:
    if os.path.isfile(OUT_FILE_NAME):
        return
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "data"
    sheet['A1'].value = 'Web'
    # sheet['B1'].value = 'ИНН'
    # sheet['C1'].value = 'Рабочий телефон (компания)'
    # sheet['D1'].value = 'Рабочий email (компания)'
    book.save(OUT_FILE_NAME)
    book.close()


def writing_data(d_info=None) -> None:
    book = openpyxl.load_workbook(filename=OUT_FILE_NAME)
    sheet = book.active
    if bool(d_info):
        sheet.append(["\n".join(d_info), ])
    else:
        sheet.append([" ", ])
    book.save(OUT_FILE_NAME)
    book.close()


def main():
    # for elem in iter(searcher()):
    #     elem = filter_links(links=elem)
    #     if len(elem) > 1:
    #         elem = duplicate_filter(links=elem)
    #     create_file()
    #     writing_data(elem)
    # create_file()
    for elem in iter(get_info_from_rusprofile()):
        print(elem)
        writing_data(elem)
        time.sleep(1)


if __name__ == '__main__':
    main()
    # check_point()
    # Список некорректных запросов:
    # //*[@id="wrapper"]/header/div/div[2]/div/div/a[1]
    # list_inn_error = ['5048058640', '7701039070', '5032322433']
