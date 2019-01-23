#!/usr/bin/env python

from time import sleep
from datetime import datetime
import re
import os
import sys
import glob
import logging
from collections import OrderedDict
from xml.etree import ElementTree as ET

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException, TimeoutException, ElementNotVisibleException)
from selenium.webdriver.remote.remote_connection import LOGGER
import xlrd
import xlwt
import xlutils.copy
import choice

import arial10

try:
    import glob2
except ImportError:
    glob2 = None


DEBUG = ('--debug' in sys.argv)
logging.basicConfig(level=(logging.DEBUG if DEBUG else logging.INFO),
                    format='%(asctime)s %(levelname)s %(message)s')

LOGGER.setLevel(logging.WARNING)

log = logging.getLogger('sfs')


def get_relative_path(path):
    dirname = os.path.dirname(__file__)
    return os.path.abspath(os.path.join(dirname, path))


def maybe_remove(path):
    if os.path.exists(path):
        os.remove(path)


WAIT_TIMEOUT = 15
LAST_REPORT_STATUS_YEAR = None  # for current year
LAST_REPORT_STATUS_YEAR = 2018  # delete this row for current year

KEY_PASSWORD = open(get_relative_path('key_password')).read().strip()
KEYS_FILENAME = get_relative_path('keys.xls')

INFO_FILENAME = get_relative_path('info.xls')
REPORT_STATUS_FILENAME = get_relative_path('report_status.xls')

KEYS_DIR = get_relative_path('./keys')

REPORTS_DIR = get_relative_path('./reports')
OUTBOX_DIR = get_relative_path('./outbox')
SENT_DIR = get_relative_path('./sent')

BUDGET_STATUS_CODES = OrderedDict((
    ('18050400', 'ЄП'),
    ('18050401', 'ЄП'),
    ('71040000', 'ЄСВ'),
    ('11011000', 'ВЗ'),
    ('11010500', 'ПДФО'),
    ('18010200', 'Недвиж'),
    ('18010700', 'Земля'),
    ('_', 'UNKNOWN'),
))


class SeleniumHelperMixin:
    def create_driver(self):
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--lang=en-US')
        chrome_options.add_experimental_option('prefs', {
            'download.default_directory': self.reports_dir,
            'safebrowsing.enabled': True,
        })
        driver = webdriver.Chrome(chrome_options=chrome_options)
        driver.set_page_load_timeout(WAIT_TIMEOUT)
        # maybe move out from screen?
        # driver.set_window_position(0, 0)
        # driver.set_window_size(800, 600)
        return driver

    def get(self, url):
        log.debug('get %s', url)
        self.driver.get(url)

    def quit(self):
        self.driver.quit()

    def get_element(self, selector, wait=False):
        if wait:
            WebDriverWait(self.driver, WAIT_TIMEOUT).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, selector)),
            )
        return self.driver.find_element_by_css_selector(selector)

    def get_elements_by_text(self, text, wait=False):
        xpath = "//*[text() = '{}']".format(text)
        if wait:
            WebDriverWait(self.driver, WAIT_TIMEOUT).until(
                EC.visibility_of_element_located((By.XPATH, xpath)),
            )
        return self.driver.find_elements_by_xpath(xpath)

    def get_element_by_text(self, *args, **kwargs):
        return self.get_elements_by_text(*args, **kwargs)[0]

    def get_element_by_text_contains(self, text, wait=False):
        xpath = "//*[text()[contains(., '{}')]]".format(text)
        if wait:
            WebDriverWait(self.driver, WAIT_TIMEOUT).until(
                EC.visibility_of_element_located((By.XPATH, xpath)),
            )
        return self.driver.find_element_by_xpath(xpath)

    def wait_presence(self, selector):
        log.debug('waiting presence %s', selector)
        return WebDriverWait(self.driver, WAIT_TIMEOUT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, selector)),
        )

    def wait_invisible(self, selector):
        log.debug('waiting invisible %s', selector)
        return WebDriverWait(self.driver, WAIT_TIMEOUT).until(
            EC.invisibility_of_element_located((By.CSS_SELECTOR, selector)),
        )

    def wait_visible(self, selector):
        log.debug('waiting visible %s', selector)
        return WebDriverWait(self.driver, WAIT_TIMEOUT).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, selector)),
        )

    def wait_callback(self, callback):
        log.debug('waiting callback')
        sleeped = 0
        while not callback():
            if sleeped >= WAIT_TIMEOUT:
                raise RuntimeError('Timeout')
            sleeped += 1
            sleep(1)

    def send_keys(self, selector, keys):
        log.debug('sending keys %s', selector)
        self.driver.find_element_by_css_selector(selector).send_keys(keys)

    def click(self, selector):
        log.debug('clicking %s', selector)
        element = self.driver.find_element_by_css_selector(selector)
        element.click()
        return element

    def wait_visible_img_and_click(self, img_url):
        selector = 'img[src="{}"]'.format(img_url)
        element = self.wait_visible(selector)
        # seems onclick event may be not FUCKING loaded!!!
        # TODO: maybe add onclick event listening to selector?
        sleep(1)  # OH YEAH...
        element.click()

    def click_img_and_wait_invisible(self, img_url):
        selector = 'img[src="{}"]'.format(img_url)
        self.click(selector)
        self.wait_invisible(selector)


class Cabinet(SeleniumHelperMixin):
    inn = fio = None

    def __init__(self, driver=None):
        self.reports_dir = REPORTS_DIR
        self.outbox_dir = OUTBOX_DIR
        self.sent_dir = SENT_DIR
        self.budget_status_report_default_path = os.path.join(self.reports_dir, 'pa.xlsx')

        for dir_ in [self.reports_dir, self.outbox_dir, self.sent_dir]:
            if not os.path.exists(dir_):
                os.mkdir(dir_)

        self.driver = driver or self.create_driver()

    def enter_cert(self, cert_path, password=KEY_PASSWORD):
        ca_map = {
            '.jks': 'АЦСК АТ КБ «ПРИВАТБАНК»',
            '.zs2': 'АЦСК ТОВ "Центр сертифікації ключів "Україна"',
        }
        for ext, ca_name in ca_map.items():
            if cert_path.lower().endswith(ext):
                select_ca = Select(self.get_element('#selectedCAs111'))
                select_ca.select_by_visible_text(ca_name)

        for pwd_filename in [cert_path + '.txt', cert_path[:cert_path.rfind('.')] + '.txt']:
            if os.path.exists(pwd_filename):
                password = open(pwd_filename).read().strip()
                break
        else:
            txt_files = tuple(glob.iglob(cert_path[:cert_path.rfind('/')] + '/*.txt'))
            if len(txt_files) == 1:
                password = open(txt_files[0]).read().strip()
                if not password:
                    # Treat filename as password if empty
                    password = os.path.basename(txt_files[0])[:-4]

        log.info('Entering cert path=%s password=%s', cert_path, password)

        cert_input = self.driver.find_elements_by_css_selector('#PKeyFileInput')[-1]
        cert_input.send_keys(cert_path)
        pwd_input = self.driver.find_elements_by_css_selector('input[type=password]')[-1]
        pwd_input.send_keys(password)
        # sleep(1)  # seems onclick is not always binded
        # self.click('#PKeyReadButton')
        self.get_elements_by_text('Зчитати')[-1].click()
        info = self.wait_visible('#certInfo').text
        match = re.match('Власник: (.+) \((\d+)\)$.*\d{2}\.\d{2}\.\d{4}-(\d{2}\.\d{2}\.\d{4})$',
                         info, re.MULTILINE | re.DOTALL)
        if not match:
            raise RuntimeError('Unmatched certInfo text: {}'.format(info))
        fio, inn, expires = match.groups()
        log.debug('inn=%s, fio=%s expires=%s', inn, fio, expires)
        return int(inn), fio, expires

    def pre_login_cert(self, cert_path, password=KEY_PASSWORD):

        # self.get('https://cabinet.sfs.gov.ua/cabinet/faces/login.jspx')
        self.get('https://cabinet.sfs.gov.ua/login')

        # self.wait_presence('.blockUI.blockOverlay')
        self.wait_invisible('.blockUI.blockOverlay')  # TODO: already renamed?
        self.wait_invisible('.ui-blockui-document')

        return self.enter_cert(cert_path, password)

    def login(self, key_path, password=KEY_PASSWORD):
        self.inn, self.fio, _ = self.pre_login_cert(key_path, password)

        login = self.driver.find_elements_by_css_selector('button[title=Увійти]')[-1]
        login.click()
        sleep(0.2)
        try:
            self.wait_invisible('.ui-blockui-document')
        except TimeoutException:
            if self.driver.current_url != 'https://cabinet.sfs.gov.ua/':
                raise
            # in other case we wasn't waiting because of redirect

        log.info('logged in inn=%s fio=%s', self.inn, self.fio)
        sleep(2)  # sleeping after login to wait redirect to new page before new get
        # self.driver.execute_script("window.stop()")  # now working

    def get_payer_info(self):
        self.get('https://cabinet.sfs.gov.ua/account')
        self.wait_visible('p-accordiontab')
        rv = OrderedDict()

        for group in self.driver.find_elements_by_css_selector('p-accordiontab'):
            group_name = group.find_element_by_css_selector('.ui-accordion-header').text.strip()
            if group_name == 'Відомості з Реєстру осіб, які здійснюють операції з товаром':
                label_postfix = ' =Товари'
            elif group_name == 'Дані про реєстрацію платником ЄСВ':
                label_postfix = ' =ЄСВ'
            else:
                label_postfix = None

            for row in group.find_elements_by_css_selector('div.row.ng-star-inserted'):
                label, value = [e.text.strip() for e in row.find_elements_by_css_selector('label')]
                if label_postfix:
                    label += label_postfix
                assert label not in rv, 'Duplicate field in "{}": "{}"'.format(group_name, label)
                rv.update({label: value})

        return rv

    def get_budget_status(self, odfs=None):
        rv = {}
        for i, data in enumerate(tuple(self.get_budget_status_items())):
            pay = data['Платіж'].split(' ')[1]
            if pay in ('18050400', '18050401', '71040000') and odfs and odfs != data['ОДФС']:
                continue
            data['saldo'] = self.get_budget_status_saldo(data, i)
            rv[BUDGET_STATUS_CODES.get(pay, tuple(BUDGET_STATUS_CODES.values())[-1])] = data
        return rv

    def _open_budget_status_page(self):
        self.get('https://cabinet.sfs.gov.ua/tax-account')
        self.wait_visible('div.ui-datalist-content')
        # self.wait_invisible('.ui-blockui-document')  # NOTE: not sure about
        try:
            self.wait_visible('div.row.data-item')
            return True
        except TimeoutException:
            return False

    def get_budget_status_items(self):
        if not self._open_budget_status_page():
            return
        for group in self.driver.find_elements_by_css_selector('div.row.data-item'):
            data = OrderedDict()
            for row in group.find_elements_by_css_selector('div.row'):
                if row.text.strip().startswith('Сплатити'):
                    continue
                try:
                    label = row.find_element_by_css_selector('label').text.strip()
                    if not label:
                        continue
                except NoSuchElementException:
                    # Seems like this is payment buttons
                    continue
                try:
                    value = row.find_element_by_css_selector('span').text.strip()
                except NoSuchElementException:
                    value = row.find_element_by_css_selector('.sum').text.strip()
                assert label not in data, 'Duplicate: {}'.format(label)
                data.update({label: value})
            yield data

    def _parse_budget_status_report_saldo(self, filename):
        wb = xlrd.open_workbook(filename)
        ws = wb.sheet_by_index(0)
        # status_date_text = ws.row(3)[1].value
        assert ws.row(4)[6].value == 'Сальдо розрахунків', 'Unexpected report format'
        try:
            saldo = ws.row(5)[6].value or 0
        except IndexError:
            saldo = 0
        # wb.close()
        return saldo

    def get_budget_status_saldo(self, data, index):
        assert self._open_budget_status_page()
        # Actually it's not safe at all, because data may have changed since we got index,
        # but fuck it
        group = self.driver.find_elements_by_css_selector('div.row.data-item')[index]
        group.click()

        self.wait_visible('i.fa-file-excel-o')
        self.wait_invisible('i.fa-spin')
        self.wait_invisible('.ui-table-loading')

        # For some reason excel table show wrong results some time, so get this from interface
        tds = self.driver.find_elements_by_css_selector('div.patable.ui-table table td')
        if tds:
            return tds[6].text or 0
        else:
            return 0

        # This is old functionality
        # maybe_remove(self.budget_status_report_default_path)
        # self.get_element('i.fa-file-excel-o').click()
        # def check_path():
        #     return os.path.exists(self.budget_status_report_default_path)
        # self.wait_callback(check_path)

        # filename = str(self.inn) + '_' + data['Платіж'].replace(' ', '') + '.xlsx'
        # filename = os.path.join(self.reports_dir, filename)
        # maybe_remove(filename)
        # os.rename(self.budget_status_report_default_path, filename)

        # return self._parse_budget_status_report_saldo(filename)

    def get_info(self):
        rv = self.get_payer_info()
        odfs = rv['Найменування ДПІ за основним місцем обліку']
        for code, data in self.get_budget_status(odfs).items():
            for k, v in data.items():
                rv[k + ' =' + code] = v
        return rv

    def get_last_report_status(self):
        self.get('https://cabinet.sfs.gov.ua/vreporting')
        report_type = self.get_element('.sticky-top .col-lg-12 .ui-dropdown-trigger', wait=True)

        if LAST_REPORT_STATUS_YEAR:
            year = self.driver.find_element_by_css_selector('input.ui-inputtext[type=text][size="10"]')
            year.clear()
            year.send_keys(str(LAST_REPORT_STATUS_YEAR))

        report_type.click()
        menu = self.driver.find_element_by_css_selector('ul.ui-dropdown-items')
        menu.find_element_by_xpath("./li/span[text() = '{}']".format('Всі')).click()
        self.wait_invisible('ul.ui-dropdown-items')
        self.wait_invisible('i.fa-spin.fa-circle-o-notch')
        headers = [td.text for td in self.driver.find_elements_by_css_selector('thead tr th')]
        assert headers[-1] == ''
        headers[-1] = 'Comment'

        values = [td.text for td in
                  self.driver.find_elements_by_css_selector('tbody tr:nth-child(1) td')]
        rv = dict(zip(headers, values))
        return rv

    def _send_report_create_form(self, code, period=None, year=None):
        code = code.upper()

        self.get('https://cabinet.sfs.gov.ua/reporting/doc/new')
        self.wait_visible('div.ui-panel-content.ui-widget-content')
        panel = self.get_element('div.ui-panel-content.ui-widget-content')

        if year:
            year_input = panel.find_elements_by_css_selector('input')[0]
            year_input.clear()
            year_input.send_keys(str(year))
            sleep(1)  # TODO: waiting table to reload

        if period:
            period = period.replace('і', 'i')  # ukrainian to english.. HAHA....
            panel.find_elements_by_css_selector('.ui-dropdown-label')[0].click()
            menu = panel.find_element_by_css_selector('ul.ui-dropdown-items')
            menu.find_element_by_xpath("./li/span[text() = '{}']".format(period)).click()
            self.wait_invisible('ul.ui-dropdown-items')
            sleep(1)  # TODO: waiting table to reload

        panel.find_elements_by_css_selector('.ui-dropdown-label')[1].click()
        menu = panel.find_element_by_css_selector('ul.ui-dropdown-items')
        for e in menu.find_elements_by_xpath("./li/span"):
            if e.text.strip().startswith(code[:3]):
                e.click()
                break
        else:
            raise ValueError('Not found type for code {}'.format(code))

        sleep(1)  # TODO: waiting table to reload
        self.get_element_by_text(code, wait=True).click()
        self.wait_visible('button i.fa.fa-plus')
        sleep(1)  # we need to fill default fields, or get error otherwise
        self.get_element('button i.fa.fa-plus').click()
        try:
            self.wait_visible('button i.fa.fa-upload')
        except TimeoutException:
            self.get_element('button i.fa.fa-plus').click()
            self.wait_visible('button i.fa.fa-upload')

    def _send_report_upload(self, filename):
        # def wait():
        #     self.wait_invisible('p-progressbar')
        #     self.wait_invisible('div[role=progressbar]')
        #     self.wait_invisible('.ui-progressbar-value')
        self.wait_visible('button i.fa.fa-upload')
        sleep(5)  # TODO: we have loaders after button appear
        file_input = self.driver.find_elements_by_css_selector('input[type="file"]')[-1]
        file_input.send_keys(filename)
        sleep(5)
        self.wait_invisible('p-progressbar')
        self.get_element('button i.fa.fa-check').click()
        self.wait_invisible('p-progressbar')
        self.get_element('button i.fa.fa-save').click()
        self.wait_invisible('p-progressbar')
        self.wait_visible('button i.fa.fa-key')

    def _send_report_sign_and_send(self, code, key_path, password):
        def _get_last(wait_icon, click=False):
            # just checking that last is the one
            assert (self.get_element('thead tr:nth-child(1)').text ==
                    'Квитанція\nСтатус\nФорма\nДата Назва')

            assert self.get_element('tbody tr:nth-child(1) td:nth-child(4)').text == code
            self.wait_visible('tbody tr:nth-child(1) i.fa.fa-{}'.format(wait_icon))
            if click:
                self.get_element('tbody tr:nth-child(1)').click()

        _get_last('check', click=True)

        self.get_element('button i.fa.fa-key').click()
        self.get_element_by_text_contains('Підпис документа', wait=True)
        inn, fio, expires = self.enter_cert(key_path, password)
        assert inn == self.inn
        assert fio == self.fio

        sign = self.driver.find_elements_by_css_selector('button[title=Підписати]')[-1]
        sign.click()
        sleep(0.2)

        _get_last('key', click=True)

        self.get_element('button i.fa.fa-send').click()

        _get_last('paper-plane', click=False)

    def send_f0103306_report(self, filename, key_path, password=KEY_PASSWORD):
        content = open(filename, 'rb').read()

        match = re.search(b'<PERIOD_YEAR>(\d+)</PERIOD_YEAR>', content, re.MULTILINE)
        assert match, 'Could not find PERIOD_YEAR (skipping) %s' % filename
        year = int(match.group(1))

        match = re.search(b'<PERIOD_MONTH>(\d+)</PERIOD_MONTH>', content, re.MULTILINE)
        assert match, 'Could not find PERIOD_MONTH (skipping) %s' % filename
        period_month = int(match.group(1))
        assert period_month in (3, 6, 9, 12), 'Unknown PERIOD_MONTH: {}'.format(period_month)
        period = {
            3: 'I квартал',
            6: 'Півріччя',
            9: '9 місяців',
            12: 'Рік',
        }[period_month]
        self._send_report_create_form('F0103306', period, year)
        self._send_report_upload(filename)
        self._send_report_sign_and_send('F0103306', key_path, password)

    def send_f3000511_report(self, filename, key_path, password=KEY_PASSWORD):
        raise NotImplementedError('F30 Not implemeneted!!! (need refactoring for new cabinet)')
        # content = open(filename, 'rb').read()
        # subreports = re.findall(b'<FILENAME>([\w\d\.]+)</FILENAME>', content)
        # subreports = [f.decode() for f in subreports]
        # if not all(os.path.exists(os.path.join(self.outbox_dir, f)) for f in subreports):
        #     raise AssertionError('Not all files exists in outbox dir: {}'.format(subreports))
        # assert len(subreports), 'No subreports found'
        # assert len(subreports) <= 2, 'Found subreports more than possible: {}'.format(subreports)

        # def create_filename(filename):
        #     return os.path.abspath(os.path.join(self.outbox_dir, filename))

        # subreports_map = {}
        # for filename_ in subreports:
        #     if 'F30501' in filename_:
        #         subreports_map['F3050111'] = create_filename(filename_)
        #     elif 'F30502' in filename_:
        #         subreports_map['F3050211'] = create_filename(filename_)
        #     else:
        #         raise AssertionError('Unknown subreport type: {}'.format(filename_))

        # self.get('https://cabinet.sfs.gov.ua/cabinet/faces/pages/dp00.jspx')
        # self.wait_connected()
        # self.wait_visible_img_and_click('/cabinet/faces/javax.faces.resource/ic_note_add.png'
        #                                 '?ln=images')
        # self.wait_connected()

        # label = self.get_element_by_text('Тип форми')
        # assert label.tag_name == 'label'
        # select = Select(label.find_element_by_xpath('../../td/select'))
        # select.select_by_value('13')  # this means F30, improve it in YOUR free time

        # self.wait_connected()
        # self.get_element_by_text('F3000511', wait=True).click()
        # self.wait_connected()

        # for k in subreports_map:
        #     (self.get_element_by_text(k)
        #      .find_element_by_xpath('../div').click())

        # sleep(1)  # just in case...
        # self.get_element_by_text('Створити ').click()
        # self.wait_connected()

        # self._send_report_upload(filename)

        # for report_name, filename in subreports_map.items():
        #     self.get_element_by_text(report_name).click()
        #     self.wait_connected()
        #     self._send_report_upload(filename)

        # self.wait_connected()
        # self._send_report_verify_sign_send(key_path, password, strict_verify=False)
        # return list(subreports_map.values())

        # # sleep(1)
        # # self.wait_connected()
        # # try:
        # #     self.get_element_by_text('OK').click()
        # #     self.wait_connected()
        # # except (NoSuchElementException, ElementNotVisibleException):
        # #     pass
        # # sleep(1)
        # # self.wait_connected()
        # # self.get_element_by_text( 'Завантажено успішно', wait=True)
        # # # e = self.get_element('.ui-pnotify-container')
        # # # assert e.text == 'Завантажено успішно', e.text


# FitSheetWrapper from https://stackoverflow.com/a/9137934/450103
class SheetWrapper(object):
    def __init__(self, sheet):
        self.sheet = sheet
        self.widths = dict()

    def write(self, r, c, label='', style=None):
        if isinstance(label, datetime) and not style:
            style = xlwt.XFStyle()
            style.num_format_str = 'YYYY-MM-DD hh:mm:ss'

        if style:
            self.sheet.write(r, c, label, style)
        else:
            self.sheet.write(r, c, label)
        width = arial10.fitwidth(style and style.num_format_str or str(label))
        if width > self.widths.get(c, 0):
            self.widths[c] = width
            self.sheet.col(c).width = int(width)

    def __getattr__(self, attr):
        return getattr(self.sheet, attr)


def open_xls(filename, headers=[]):
    if os.path.exists(filename):
        rb = xlrd.open_workbook(filename, formatting_info=True)
        wb = xlutils.copy.copy(rb)
        ws = SheetWrapper(wb.get_sheet(0))
    else:
        wb = xlwt.Workbook()
        ws = SheetWrapper(wb.add_sheet('0'))
        for i, header in enumerate(headers):
            ws.write(0, i, header)
    return ws, wb


def append_xls(filename, headers, *rows):
    ws, wb = open_xls(filename, headers)
    for row in rows:
        y = len(ws.rows)
        for x, cell in enumerate(row):
            ws.write(y, x, cell)
    wb.save(filename)


def write_row_by_index_xls(filename, index, row):
    ws, wb = open_xls(filename, None)
    for x, cell in enumerate(row):
        ws.write(index, x, cell)
    wb.save(filename)


class KeysMap(dict):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.load()

    def load(self, filename=KEYS_FILENAME):
        if os.path.exists(filename):
            wb = xlrd.open_workbook(filename, formatting_info=False)
            ws = wb.sheet_by_index(0)
            for i in range(1, ws.nrows):
                row = ws.row(i)
                inn = int(row[0].value)
                path = row[2].value
                self[inn] = path
            # wb.close()

        skipped = []
        for inn, path in self.items():
            if not os.path.exists(path):
                log.warning('Skipping %s on key load: file not found %s', inn, path)
                skipped.append(inn)

        for inn in skipped:
            self.pop(inn)

        log.info('Keys loaded (%s) from %s', len(self), filename)

    def add_key(self, inn, fio, filename, expires):
        self[inn] = filename
        append_xls(get_relative_path('keys.xls'), ['inn', 'fio', 'filename', 'expires'],
                   [inn, fio, filename, expires])


def scan_keys(keys_dir=KEYS_DIR):
    keys_map = KeysMap()
    patterns = ['Key-6.dat', '*.jks', '*.zs2', '*.ZS2']
    if glob2:
        files = sum((list(glob2.iglob(os.path.join(KEYS_DIR, '**/%s' % pattern)))
                    for pattern in patterns), [])
    else:
        files = sum((list(glob.iglob(os.path.join(KEYS_DIR, '**/%s' % pattern), recursive=True))
                    for pattern in patterns), [])

    log.info('Keys (%s) in %s', len(files), keys_dir)
    for filename in files:
        filename = os.path.abspath(filename)
        if filename not in keys_map.values():
            log.info('Checking new key %s', filename)
            cabinet = Cabinet()
            try:
                inn, fio, expires = cabinet.pre_login_cert(filename)
            except Exception as e:
                log.exception('Error occured on key processing %s %s', filename, repr(e))
                if DEBUG:
                    import pdb; pdb.set_trace()  # noqa
                continue
            finally:
                cabinet.quit()

            log.info('Adding key inn=%s fio=%s expires=%s filename=%s',
                     inn, fio, expires, filename)
            keys_map.add_key(inn, fio, filename, expires)


def _get_report(filename, headers, method_name, method_kwargs={}):
    log.info('Populating report %s', filename)

    keys_map = KeysMap()
    processed = []

    if os.path.exists(filename):
        wb = xlrd.open_workbook(filename)
        ws = wb.sheet_by_index(0)
        for i in range(1, ws.nrows):
            inn = ws.row(i)[0].value
            if inn:
                processed.append(int(inn))

    to_process = set(keys_map) - set(processed)
    log.info('Processing %s (processed already %s)', len(to_process), len(set(processed)))

    for inn in to_process:
        cabinet = Cabinet()
        try:
            cabinet.login(keys_map[inn])
            assert cabinet.inn == inn, 'Key inn in store and after login not matched!'
            data = getattr(cabinet, method_name)(**method_kwargs)
        except Exception as e:
            log.exception('Error occured on %s processing %s %s', method_name, inn, repr(e))
            if DEBUG:
                import pdb; pdb.set_trace()  # noqa
            continue
        finally:
            cabinet.quit()
        log.info('Adding row inn=%s fio=%s data=%s',
                 cabinet.inn, cabinet.fio, data)
        skipped_headers = [h for h in headers if h not in data]
        if skipped_headers:
            log.warning('Skipped headers: %s',  skipped_headers)
        row = [data.get(k, '') for k in headers]
        append_xls(filename,
                   ['inn', 'fio', 'parsed'] + headers,
                   [cabinet.inn, cabinet.fio, datetime.now()] + row)


def get_info(filename=INFO_FILENAME):
    headers = [
        'Прізвище, ім’я та по батькові',
        'Податковий номер',
        'Особливий режим',
        'Телефони',
        'Дата зняття з обліку',
        'Номер взяття на облік платника податків',
        'Найменування ДПІ за основним місцем обліку',
        'Код ДПІ за основним місцем обліку',
        'Адреса',
        'Дата взяття на облік платника податків',
        'Дата реєстрації платником податку',
        'Дата анулювання реєстрації',
        'Термін дії реєстрації',
        'Підстава анулювання',
        'Причина анулювання',
        'Індивідуальний податковий номер',
        'Група',
        'Дата анулювання',
        'Ставка',
        'Дата переходу на спрощену систему оподаткування',
        'Дата взяття на облік =ЄСВ',
        'Клас професійного ризику виробництва =ЄСВ',
        'Код КВЕД по якому призначено клас професійного ризику =ЄСВ',
        'Дата зняття з обліку =ЄСВ',
        'Реєстраційний номер платника єдиного внеску =ЄСВ',
        'Обліковий номер особи =Товари',
        'Дата взяття на облік =Товари',
        'Дата зняття з обліку =Товари',
        'Дата внесення змін =Товари',
    ]
    for k in OrderedDict((k, None) for k in BUDGET_STATUS_CODES.values()).keys():
        headers.extend([
            'ОДФС =' + k,
            'Назва податку =' + k,
            'Платіж =' + k,
            'Код ЄДРПОУ отримувача =' + k,
            'МФО =' + k,
            'Назва отримувача =' + k,
            'Бюджетний рахунок =' + k,
            'Нараховано/зменшено =' + k,
            'Сплачено до бюджету =' + k,
            'Повернуто з бюджету =' + k,
            'Пеня =' + k,
            'Недоїмка =' + k,
            'Переплата =' + k,
            'Залишок несплаченої пені =' + k,
            'saldo =' + k,
        ])

    _get_report(filename, headers, 'get_info')


def get_report_status(filename=REPORT_STATUS_FILENAME):
    headers = ['ДФС', 'Форма', 'Номер', 'Дата', 'Період', 'Додатки', 'Comment']
    _get_report(filename, headers, 'get_last_report_status')


def send_outbox(outbox_dir=OUTBOX_DIR, sent_dir=SENT_DIR):
    keys_map = KeysMap()
    files = tuple(glob.iglob(os.path.join(OUTBOX_DIR, '*.xml')))
    log.info('Outbox (%s) in %s', len(files), outbox_dir)

    for filename in files:
        filename = os.path.abspath(filename)
        if not os.path.exists(filename):
            continue  # this was subreport and was processed already
        content = open(filename, 'rb').read()

        match = re.search(b'<TIN>(\d+)</TIN>', content, re.MULTILINE)
        if not match:
            log.error('Could not find inn (skipping) %s', filename)
            continue

        inn = int(match.group(1))
        if inn not in keys_map:
            log.error('inn %s not found in keys map (skipping) %s', inn, filename)
            continue

        match = re.search(b'<C_DOC>([\w\d]+)</C_DOC>', content, re.MULTILINE)
        if not match:
            log.error('%s: report type not found (skipping) %s', inn, filename)
            continue
        report_type = match.group(1).decode().upper()
        assert report_type in ['F30', 'F01'], report_type

        if report_type == 'F30':
            match = re.search(b'<C_DOC_SUB>(\d+)</C_DOC_SUB>', content, re.MULTILINE)
            if int(match.group(1)) != 5:
                continue  # this is subreport, so processing only head report

        cabinet = Cabinet()
        try:
            cabinet.login(keys_map[inn])
            assert cabinet.inn == inn, 'Key inn in store and after login not matched!'
            if report_type == 'F01':
                subreports = cabinet.send_f0103306_report(filename, key_path=keys_map[inn])
            elif report_type == 'F30':
                subreports = cabinet.send_f3000511_report(filename, key_path=keys_map[inn])
            else:
                raise AssertionError('Unknown report type: {}'.format(report_type))
        except Exception as e:
            log.exception('Error occured on outbox processing %s %s', filename, repr(e))
            if DEBUG:
                import pdb; pdb.set_trace()  # noqa
            continue
        finally:
            cabinet.quit()

        log.info('Sent report inn=%s fio=%s filename=%s', cabinet.inn, cabinet.fio,
                 os.path.basename(filename))
        for filename in ([filename] + (subreports and list(subreports) or [])):
            dest = os.path.join(sent_dir, os.path.basename(filename))
            maybe_remove(dest)
            os.rename(filename, dest)
            # TODO: add to xls log, or rename budget_status to just status?


if __name__ == '__main__':
    funcs = ['scan_keys', 'get_info', 'get_report_status', 'send_outbox']
    try:
        func = choice.Menu(funcs).ask()
        globals()[func]()
    except Exception as e:
        log.exception(repr(e))
    input('DONE. press any key to close')
