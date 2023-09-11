import os
import time
import json
import shutil

from datetime import date
from datetime import timedelta
from dateutil.relativedelta import relativedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service

from loguru import logger

reports_path = os.path.join(os.path.abspath(os.getcwd()), 'reports')

options = webdriver.ChromeOptions()
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_argument("--start-maximized")
options.add_argument("--disable-extensions")
options.add_argument("--disable-popup-blocking")
options.add_argument("--headless=new")
options.add_experimental_option("prefs", {
  "download.default_directory": reports_path,
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})

service = Service('C:\chromedriver\chromedriver.exe')
browser = webdriver.Chrome(options=options, service=service)

actions = ActionChains(browser)

def retry(exception=Exception, retries=5, delay=0):
    def decorator(func):
        def wrapper(*args, **kwargs):
            for i in range(retries):
                try:
                    return func(*args, **kwargs)
                except exception as ex:
                    logger.exception(ex)
                    logger.debug(f'Попытка выполнить {func.__name__}: {i}/{retries}')
                    time.sleep(delay)
            raise ex
        return wrapper
    return decorator

def download_wait(directory, timeout, nfiles=None):
    """
    Wait for downloads to finish with a specified timeout.

    Args
    ----
    directory : str
        The path to the folder where the files will be downloaded.
    timeout : int
        How many seconds to wait until timing out.
    nfiles : int, defaults to None
        If provided, also wait for the expected number of files.

    """
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < timeout:
        time.sleep(1)
        dl_wait = False
        files = os.listdir(directory)
        if nfiles and len(files) != nfiles:
            dl_wait = True

        for fname in files:
            if fname.endswith('.crdownload'):
                dl_wait = True

        seconds += 1
    return seconds

def autorization(login_data: str, password_data: str):
    logger.debug('Начата авторизация')

    browser.get('http://bi.mz.mosreg.ru/login/')

    login_field = browser.find_element(By.XPATH, '//*[@id="login"]')
    login_field.send_keys(login_data)

    password_field = browser.find_element(By.XPATH, '//*[@id="password"]')
    password_field.send_keys(password_data)

    browser.find_element(By.XPATH, '//*[@id="isLoginBinding"]/form/div[4]/button').click()

    logger.debug('Авторизация пройдена')


def open_bi_report(report_name, begin_date, end_date):
    logger.debug(f'Открываю {report_name} - выбран период: с {begin_date.strftime("%d.%m.%Y")} по {end_date.strftime("%d.%m.%Y")}')
    browser.get('http://bi.mz.mosreg.ru/#form/' + report_name)
    WebDriverWait(browser, 60).until(EC.element_to_be_clickable((By.XPATH, "//input[@data-componentid='ext-datefield-3']")))
    browser.execute_script("var first_date = globalThis.Ext.getCmp('ext-datefield-3'); +\
                           first_date.setValue('" + begin_date.strftime('%d.%m.%Y') + "'); + \
                           first_date.fireEvent('select');")
    browser.execute_script("var last_date = globalThis.Ext.getCmp('ext-datefield-4'); +\
                           last_date.setValue('" + end_date.strftime('%d.%m.%Y') + "'); + \
                           last_date.fireEvent('select');")
          
    # Смотрим сколько загружено записей
    #count_values = int(browser.execute_script("return globalThis.Ext.getCmp('ext-numberfield-1').getValue();"))
    # Если загружено 0 ожидаем загрузки
    #while count_values == 0:
    #    time.sleep(5)
    #    count_values = int(browser.execute_script("return globalThis.Ext.getCmp('ext-numberfield-1').getValue();"))
    # Нажать на кнопку "Обновить" для загрузки отчёта по нашим датам
    
    WebDriverWait(browser, 300).until(EC.invisibility_of_element((By.XPATH, '//div[@data-componentid="ext-toolbar-8"]')))

    # Фильтр ОГРН
    if report_name == 'pass_dvn':
        #element = browser.find_element(By.XPATH, '//*[@id="ext-RTA-gridview-1"]/div[1]/div/table/tbody/tr[2]/td[3]/div/div/div[1]/div[1]/div[1]')
        #ActionChains(browser).click(element).send_keys('1215000036305').perform()
        browser.execute_script("var ogrn_filter = globalThis.Ext.getCmp('ext-RTA-grid-textfilter-35'); +\
                            ogrn_filter.setValue('1215000036305'); + \
                            ogrn_filter.fireEvent('select');")

        #element = browser.find_element(By.XPATH, '/html/body/div[3]/div[4]/div/div[1]/div/div/div[1]/div[3]/div[1]/div[2]/div/div/div[1]/div/div')
        #element.click()

    browser.find_element(By.XPATH, "//button[@data-componentid='ext-button-12']").click()

    WebDriverWait(browser, 300).until(EC.invisibility_of_element((By.XPATH, '//div[@data-componentid="ext-toolbar-8"]')))

    # Смотрим сколько загружено записей
    #new_count_values = int(browser.execute_script("return globalThis.Ext.getCmp('ext-numberfield-1').getValue();"))
    # Если загружено, столько же, сколько в начале, то ожидаем формирования отчёта по нашим датам
    #while new_count_values == count_values:
    #    time.sleep(5)
    #    new_count_values = int(browser.execute_script("return globalThis.Ext.getCmp('ext-numberfield-1').getValue();"))
    #    logger.debug(f'{new_count_values} == {count_values}')
    #logger.debug('Отчет Прохождение пациентами ДВН или ПМО загружен в браузере')   
    
def save_report():
    logger.debug('Начинается сохранение файла с отчетом в папку: ' + reports_path)
    try:
        os.mkdir(reports_path)
    except FileExistsError:
        pass

    # Нажимаем на кнопку "Выгрузить в Excel" и ожидаем загрузку файла
    browser.find_element(By.XPATH, "//button[@data-componentid='ext-button-13']").click()

    download_wait(reports_path, 600, len(os.listdir(reports_path)) + 1)

    browser.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/div/div/div[2]/div/div/div[1]/div[1]/div[2]/div/div[3]/div[4]').click()
    
    logger.debug('Сохранение файла с отчетом успешно')

def start_report_saving():
    shutil.rmtree(reports_path, ignore_errors=True) # Очистить предыдущие результаты
    credentials_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'auth-bi-emias.json')

    f = open(credentials_path, 'r', encoding='utf-8')
    data = json.load(f)
    f.close()
    for _departments in data['departments']:        
        for _units in _departments["units"]:
            autorization(_units['login'], _units['password'])

    # Установка дат
    # С начала недели
    first_date = date.today() - timedelta(days=date.today().weekday()) # начало текущей недели
    last_date = date.today() # сегодня
 
    # Если сегодня понедельник, то берем всю прошлую неделю
    if date.today() == (date.today() - timedelta(days=date.today().weekday())):
        first_date = date.today() - timedelta(days=date.today().weekday()) - timedelta(days=7) # начало прошлой недели

    open_bi_report('disp_tmk', first_date, last_date)
    save_report()

    # Установка дат
    #first_date = date.today() - relativedelta(months=1) - timedelta(days=1) # минус один месяц
    #last_date = date.today() # сегодня

    #browser.refresh()

    #open_bi_report('pass_dvn', first_date, last_date)
    #save_report()
    
    logger.debug('Выгрузка из BI ЕМИАС завершена')

start_report_saving()

#browser.quit()