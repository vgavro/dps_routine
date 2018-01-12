#!/usr/bin/env python

from time import sleep
from datetime import datetime
import re
import os
import sys
import glob
import logging
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


logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s %(levelname)s %(message)s')

LOGGER.setLevel(logging.WARNING)

log = logging.getLogger('sfs')


def get_relative_path(path):
    dirname = os.path.dirname(__file__)
    return os.path.abspath(os.path.join(dirname, path))


def maybe_remove(path):
    if os.path.exists(path):
        os.remove(path)


WAIT_TIMEOUT = 20
DEBUG = ('--debug' in sys.argv)

KEY_PASSWORD = open(get_relative_path('key_password')).read().strip()
KEYS_FILENAME = get_relative_path('keys.xls')
PAYER_INFO_FILENAME = get_relative_path('payer_info.xls')
BUDGET_STATUS_FILENAME = get_relative_path('budget_status.xls')
RECEIPTS_STATUS_FILENAME = get_relative_path('receipts_status.xls')
KEYS_DIR = get_relative_path('./keys')

REPORTS_DIR = get_relative_path('./reports')
OUTBOX_DIR = get_relative_path('./outbox')
SENT_DIR = get_relative_path('./sent')


class Cabinet:
    inn = fio = None

    def __init__(self, driver=None):
        self.reports_dir = REPORTS_DIR
        self.outbox_dir = OUTBOX_DIR
        self.sent_dir = SENT_DIR
        self.budget_status_report_default_path = os.path.join(self.reports_dir, 'pa.xls')
        self.receipt_xml_default_path = os.path.join(self.reports_dir, 'data.xml')
        self.card_payer_default_path = os.path.join(self.reports_dir, 'card_pp.html')

        for dir_ in [self.reports_dir, self.outbox_dir, self.sent_dir]:
            if not os.path.exists(dir_):
                os.mkdir(dir_)

        if not driver:
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
            # driver.set_window_size(0, 0)
        self.driver = driver

    def get(self, url):
        log.debug('get %s', url)
        self.driver.get(url)

    def quit(self):
        self.driver.quit()

    def get_element(self, selector):
        return self.driver.find_element_by_css_selector(selector)

    def get_element_by_text(self, text, wait=False):
        xpath = "//*[text() = '{}']".format(text)
        if wait:
            WebDriverWait(self.driver, WAIT_TIMEOUT).until(
                EC.visibility_of_element_located((By.XPATH, xpath)),
            )
        return self.driver.find_element_by_xpath(xpath)

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

    def wait_connected(self):
        self.wait_visible('img[src="/cabinet/afr/alta-v1/connected.gif"]')

    def enter_cert(self, cert_path, password=KEY_PASSWORD):
        self.send_keys('#PKeyFileInput', cert_path)
        self.send_keys('#PKeyPassword', password)
        sleep(1)  # seems onclick is not always binded
        self.click('#PKeyReadButton')
        info = self.wait_visible('#certInfo').text
        match = re.match('Власник: (.+) \((\d+)\)$', info, re.MULTILINE)
        if not match:
            raise RuntimeError('Unmatched certInfo text: {}'.format(info))
        fio, inn = match.groups()
        log.debug('inn=%s, fio=%s', inn, fio)
        return int(inn), fio

    def pre_login_cert(self, cert_path, password=KEY_PASSWORD):
        self.get('https://cabinet.sfs.gov.ua/cabinet/faces/login.jspx')

        self.wait_presence('.blockUI.blockOverlay')
        self.wait_invisible('.blockUI.blockOverlay')
        return self.enter_cert(cert_path, password)

    def login(self, key_path, password=KEY_PASSWORD):
        self.inn, self.fio = self.pre_login_cert(key_path, password)
        self.click('#LoginButton')
        try:
            self.wait_connected()
        except TimeoutException:
            if self.driver.current_url != 'https://cabinet.sfs.gov.ua/':
                raise
            # in other case we wasn't waiting because of redirect

        log.info('logged in inn=%s fio=%s', self.inn, self.fio)
        sleep(2)  # sleeping after login to wait redirect to new page before new get
        # self.driver.execute_script("window.stop()")  # now working

    def get_budget_status_report(self, payment_id):
        maybe_remove(self.budget_status_report_default_path)
        # self.get('https://cabinet.sfs.gov.ua/cabinet/faces/index.jspx')
        # self.wait_connected()
        # getting first url to reset state of second one
        self.get('https://cabinet.sfs.gov.ua/cabinet/faces/pages/ta.jspx')
        self.get('https://cabinet.sfs.gov.ua/cabinet/faces/pages/ta_ext.jspx')
        # self.wait_connected()

        e = self.get_element_by_text_contains(payment_id, wait=True)
        payment_info = e.find_element_by_xpath('../../../../..').text

        assert payment_info.startswith('ОДФС'), 'Unknown payment info: {}'.format(payment_info)

        patterns = [
            '^ОДФС (.*)$',
            '^Код ЄДРПОУ отримувача (\d+)$',
            '^МФО (\d+)$',
            '^Бюджетний рахунок (\d+)$',
        ]
        payment_info_parsed = []
        for pattern in patterns:
            match = re.search(pattern, payment_info, re.MULTILINE)
            assert match, 'Pattern not matched {}: {}'.format(pattern, payment_info)
            payment_info_parsed.append(match.group(1))

        e.click()
        # self.wait_connected()

        self.wait_visible_img_and_click('/cabinet/faces/javax.faces.resource/'
                                        'microsoft-excel.png?ln=images')

        def check_path():
            return os.path.exists(self.budget_status_report_default_path)
        self.wait_callback(check_path)

        filename = str(self.inn) + '_' + payment_id.replace(' ', '') + '.xls'
        filename = os.path.join(self.reports_dir, filename)
        maybe_remove(filename)
        os.rename(self.budget_status_report_default_path, filename)
        return payment_info_parsed, filename

        # self.click_img_and_wait_invisible('/cabinet/faces/javax.faces.resource/back.png?ln=images')
        # self.wait_connected()

    def get_budget_status(self):
        def parse_saldo(filename):
            wb = xlrd.open_workbook(filename)
            ws = wb.sheet_by_index(0)
            # status_date_text = ws.row(3)[1].value
            assert ws.row(5)[6].value == 'Сальдо розрахунків', 'Unexpected report format'
            try:
                saldo = ws.row(6)[6].value or 0
            except IndexError:
                saldo = 0
            # wb.close()
            return saldo

        try:
            info1, filename = self.get_budget_status_report('18050400')
        except NoSuchElementException:
            info1, filename = self.get_budget_status_report('18050401')
        saldo1 = parse_saldo(filename)
        info1.append(saldo1)  # EN

        info2, filename = self.get_budget_status_report('71040000')
        saldo2 = parse_saldo(filename)
        info2.append(saldo2)  # ESV

        return info1, info2

    def get_last_receipt(self):
        def parse_receipt(filename):
            root = ET.parse(filename).getroot()

            year = int(root.find('DECLARHEAD/PERIOD_YEAR').text)
            month = int(root.find('DECLARHEAD/PERIOD_MONTH').text)
            sent_date = root.find('DECLARBODY/HDATE').text
            sent_time = root.find('DECLARBODY/HTIME').text
            report_type = root.find('DECLARBODY/HDOCKOD').text
            result_text = root.find('DECLARBODY/HRESULT').text
            doc_state = root.find('DECLARBODY/HDOCSTAN').text
            if result_text in ('Пакет прийнято.', 'Прийнято пакет.'):
                status = 2
            else:
                status = 1
            return report_type, status, year, month, sent_date, sent_time, result_text, doc_state

        maybe_remove(self.receipt_xml_default_path)

        self.get('https://cabinet.sfs.gov.ua/cabinet/faces/pages/dm03_ext.jspx')
        # self.wait_connected()
        # sleep(2)
        try:
            self.get_element_by_text_contains('[J1499201]').click()
        except NoSuchElementException:
            # import pdb; pdb.set_trace()
            return None
        # self.wait_connected()
        self.wait_visible_img_and_click('/cabinet/faces/javax.faces.resource/'
                                        'xml.png?ln=images')
        self.wait_callback(lambda: os.path.exists(self.receipt_xml_default_path))
        filename = str(self.inn) + '_J1499201.xls'
        filename = os.path.join(self.reports_dir, filename)
        maybe_remove(filename)
        os.rename(self.receipt_xml_default_path, filename)
        return parse_receipt(filename)

    def get_payer_info(self):
        if os.path.exists(self.card_payer_default_path):
            os.remove(self.card_payer_default_path)

        self.get('https://cabinet.sfs.gov.ua/cabinet/faces/pages/rg_ext.jspx')
        self.get_element_by_text('ПЕРЕГЛЯД ДАНИХ', wait=True).click()
        self.get_element_by_text('ДРУКУВАТИ', wait=True).click()

        self.wait_callback(lambda: os.path.exists(self.card_payer_default_path))
        filename = str(self.inn) + '_card_payer.html'
        filename = os.path.join(self.reports_dir, filename)
        maybe_remove(filename)
        os.rename(self.card_payer_default_path, filename)
        content = open(filename, 'rb').read().decode('utf8', errors='ignore')
        data = dict(re.findall('<tr><td[^>]*>([^>]+)</td><td[^>]+>([^>]+)</td></tr>', content))
        for k in data:
            if data[k] == '&nbsp':
                data[k] = ''
        return data

    def send_f3000511_report(self, filename, key_path, password=KEY_PASSWORD):
        content = open(filename, 'rb').read()
        subreports = re.findall(b'<FILENAME>([\w\d\.]+)</FILENAME>', content)
        subreports = [f.decode() for f in subreports]
        if not all(os.path.exists(os.path.join(self.outbox_dir, f)) for f in subreports):
            raise AssertionError('Not all files exists in outbox dir: {}'.format(subreports))
        assert len(subreports), 'No subreports found'
        assert len(subreports) <= 2, 'Found subreports more than possible: {}'.format(subreports)

        subreports_map = {}
        for filename_ in subreports:
            if 'F30501' in filename_:
                subreports_map['F3050111'] = os.path.abspath(os.path.join(self.outbox_dir, filename_))
            elif 'F30502' in filename_:
                subreports_map['F3050211'] = os.path.abspath(os.path.join(self.outbox_dir, filename_))
            else:
                raise AssertionError('Unknown subreport type: {}'.format(filename_))

        self.get('https://cabinet.sfs.gov.ua/cabinet/faces/pages/dp00.jspx')
        self.wait_connected()
        self.wait_visible_img_and_click('/cabinet/faces/javax.faces.resource/ic_note_add.png'
                                        '?ln=images')
        self.wait_connected()

        label = self.get_element_by_text('Тип форми')
        assert label.tag_name == 'label'
        select = Select(label.find_element_by_xpath('../../td/select'))
        select.select_by_value('13')  # this means F30, improve it in YOUR free time

        self.wait_connected()
        self.get_element_by_text('F3000511', wait=True).click()
        self.wait_connected()

        for k in subreports_map:
            (self.get_element_by_text(k)
             .find_element_by_xpath('../div').click())

        sleep(1)  # just in case...
        self.get_element_by_text('Створити ').click()
        self.wait_connected()

        self._send_report_upload(filename)

        for report_name, filename in subreports_map.items():
            self.get_element_by_text(report_name).click()
            self.wait_connected()
            self._send_report_upload(filename)

        self.wait_connected()
        self._send_report_verify_sign_send(key_path, password)
        return list(subreports_map.values())

    def _send_report_upload(self, filename):
        self.wait_visible_img_and_click('/cabinet/faces/javax.faces.resource/upload.png'
                                        '?ln=images')
        sleep(1)
        self.wait_connected()
        sleep(1)
        try:
            self.get_element_by_text('Оновити ...').click()
            sleep(1)
        except NoSuchElementException:
            pass
        self.wait_connected()
        self.send_keys('input[type="file"]', filename)
        sleep(1)
        self.wait_connected()
        try:
            self.get_element_by_text('OK').click()
            self.wait_connected()
        except (NoSuchElementException, ElementNotVisibleException):
            pass
        sleep(1)
        self.wait_connected()
        e = self.get_element('.ui-pnotify-container')
        assert e.text == 'Завантажено успішно', e.text

    def _send_report_verify_sign_send(self, key_path, password):
        self.wait_visible_img_and_click('/cabinet/faces/javax.faces.resource/checked.png?ln=images')
        sleep(1)
        self.wait_connected()
        sleep(1)  # well, you may remove this shit if you have enough time for cabinet debug...
        try:
            e = self.get_element('.ui-pnotify-container')
        except NoSuchElementException:
            raise RuntimeError('Звіт має помилки (не критичні?)')
        assert e.text == 'Помилок немає', e.text

        self.wait_visible_img_and_click('/cabinet/faces/javax.faces.resource/sign.png?ln=images')
        self.wait_connected()

        self.get_element_by_text('Підпис документа  приватним підприємцем')
        # checked that all ok

        inn, fio = self.enter_cert(key_path, password)
        assert inn == self.inn
        assert fio == self.fio

        self.click('#LoginButton')
        sleep(1)
        self.wait_connected()
        sleep(1)  # yeah... i know...
        e = self.wait_visible('.ui-pnotify-container')
        assert str(e.text).startswith('Підписано успішно')
        return e.text

    def send_f0103306_report(self, filename, key_path, password=KEY_PASSWORD):
        content = open(filename, 'rb').read()

        match = re.search(b'<PERIOD_YEAR>(\d+)</PERIOD_YEAR>', content, re.MULTILINE)
        assert match, 'Could not find PERIOD_YEAR (skipping) %s' % filename
        period_year = int(match.group(1))

        match = re.search(b'<PERIOD_MONTH>(\d+)</PERIOD_MONTH>', content, re.MULTILINE)
        assert match, 'Could not find PERIOD_MONTH (skipping) %s' % filename
        period_month = int(match.group(1))
        assert period_month in (3, 6, 9, 12), 'Unknown PERIOD_MONTH: {}'.format(period_month)

        # self.get('https://cabinet.sfs.gov.ua/cabinet/faces/index.jspx')
        # self.wait_connected()
        self.get('https://cabinet.sfs.gov.ua/cabinet/faces/pages/dp00.jspx')
        self.wait_connected()
        self.wait_visible_img_and_click('/cabinet/faces/javax.faces.resource/ic_note_add.png'
                                        '?ln=images')
        self.wait_connected()

        label = self.get_element_by_text('Рік')
        assert label.tag_name == 'label'
        input_ = label.find_element_by_xpath('../../td/table/tbody/tr/td/input')
        input_.clear()
        input_.send_keys(str(period_year))

        label = self.get_element_by_text('Період')
        assert label.tag_name == 'label'
        select = Select(label.find_element_by_xpath('../../td/select'))
        select.select_by_value(str(period_month - 1))  # months 0...11

        # this should invoke list loading, not working without it
        # TODO: click not on report, but on some safe place
        REPORT_NAME = 'Податкова декларацiя платника єдиного податку - фiзичної особи _ пiдприємця'
        self.get_element_by_text(REPORT_NAME).click()
        sleep(2)
        self.wait_connected()

        self.get_element_by_text(REPORT_NAME).click()
        self.wait_connected()

        #  Comment lines block below to leace "звітний документ" by default
        #  document_state = Select(self.get_element('select[title="звітний документ"]'))
        #  document_state.select_by_value('1')  #  новий звітній документ
        #  sleep(1)

        self.get_element_by_text('Створити ').click()
        self.wait_connected()

        self._send_report_upload(filename)
        self._send_report_verify_sign_send(key_path, password)


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
        for inn, path in self.items():
            if not os.path.exists(path):
                log.warn('Skipping %s on key load: file not found %s', inn, path)
                self.pop(inn)
        log.info('Keys loaded (%s) from %s', len(self), filename)

    def add_key(self, inn, fio, filename):
        self[inn] = filename
        append_xls(get_relative_path('keys.xls'), ['inn', 'fio', 'filename'], [inn, fio, filename])


def scan_keys(keys_dir=KEYS_DIR):
    keys_map = KeysMap()
    if glob2:
        files = tuple(glob2.iglob(os.path.join(KEYS_DIR, '**/Key-6.dat')))
    else:
        files = tuple(glob.iglob(os.path.join(KEYS_DIR, '**/Key-6.dat'),
                                 recursive=True))
    log.info('Keys (%s) in %s', len(files), keys_dir)
    for filename in files:
        filename = os.path.abspath(filename)
        if filename not in keys_map.values():
            log.info('Checking new key %s', filename)
            cabinet = Cabinet()
            try:
                inn, fio = cabinet.pre_login_cert(filename)
            except Exception as e:
                log.exception('Error occured on key processing %s %s', filename, repr(e))
                if DEBUG:
                    import pdb; pdb.set_trace()  # noqa
                continue
            finally:
                cabinet.quit()

            log.info('Adding key inn=%s fio=%s filename=%s', inn, fio, filename)
            keys_map.add_key(inn, fio, filename)


def get_budget_status(filename=BUDGET_STATUS_FILENAME):
    log.info('Get budget status %s', filename)

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
            en_info, esv_info = cabinet.get_budget_status()
        except Exception as e:
            log.exception('Error occured on budget processing %s %s', inn, repr(e))
            if DEBUG:
                import pdb; pdb.set_trace()  # noqa
            continue
        finally:
            cabinet.quit()
        log.info('Adding budget status inn=%s fio=%s en_info=%s esv_info=%s',
                 cabinet.inn, cabinet.fio, en_info, esv_info)
        append_xls(filename,
                   ['inn', 'fio', 'parsed',
                    'en_odfs', 'en_edrpou', 'en_mfo', 'en_account', 'en_saldo',
                    'esv_odfs', 'esv_edrpou', 'esv_mfo', 'esv_account', 'esv_saldo',
                   ],
                   [cabinet.inn, cabinet.fio, datetime.now(),
                    en_info[0], en_info[1], en_info[2], en_info[3], en_info[4],
                    esv_info[0], esv_info[1], esv_info[2], esv_info[3], esv_info[4],
                   ])


def get_payer_info(filename=PAYER_INFO_FILENAME):
    log.info('Get payer info %s', filename)

    headers = [
        'Податковий номер',
        'Прізвище, ім’я та по батькові',
        'Код ДПІ за основним місцем обліку',
        'Найменування ДПІ за основним місцем обліку',
        'Дата взяття на облік платника податків',
        'Номер взяття на облік платника податків',
        'Особливий режим',
        'Дата зняття з обліку',
        'Адреса',
        'Телефони',
        'Дата переходу на спрощену систему оподаткування',
        'Група',
        'Ставка',
        'Дата анулювання',
        'Дата взяття на облік',
        'Реєстраційний номер платника єдиного внеску',
        'Клас професійного ризику виробництва',
        'Код КВЕД по якому призначено клас професійного ризику',
        'Дата зняття з обліку ',
        'Код ДПІ за неосновним місцем обліку',
    ]

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
            info = cabinet.get_payer_info()
        except Exception as e:
            log.exception('Error occured on payer info processing %s %s', inn, repr(e))
            if DEBUG:
                import pdb; pdb.set_trace()  # noqa
            continue
        finally:
            cabinet.quit()
        log.info('Adding payer info inn=%s fio=%s info=%s',
                 cabinet.inn, cabinet.fio, info)
        data = [info.get(k, '') for k in headers]
        append_xls(filename,
                   ['inn', 'fio', 'parsed'] + headers,
                   [cabinet.inn, cabinet.fio, datetime.now()] + data)


def get_receipts_status(filename=RECEIPTS_STATUS_FILENAME):
    log.info('Get receipts status %s', filename)

    keys_map = KeysMap()
    processed = []
    inn_rows = {}

    if os.path.exists(filename):
        wb = xlrd.open_workbook(filename)
        ws = wb.sheet_by_index(0)
        for i in range(1, ws.nrows):
            inn = ws.row(i)[0].value
            if inn and int(ws.row(i)[4].value or 0) == 2:
                # no check if status == 2
                processed.append(int(inn))
            else:
                inn_rows[inn] = i

    to_process = set(keys_map) - set(processed)
    log.info('Processing %s (processed already %s)', len(to_process), len(set(processed)))

    for inn in to_process:
        cabinet = Cabinet()
        try:
            cabinet.login(keys_map[inn])
            assert cabinet.inn == inn, 'Key inn in store and after login not matched!'
            info = cabinet.get_last_receipt()
        except Exception as e:
            log.exception('Error occured on receipt processing %s %s', inn, repr(e))
            if DEBUG:
                import pdb; pdb.set_trace()  # noqa
            continue
        finally:
            cabinet.quit()
        log.info('Adding recepit status inn=%s fio=%s info=%s',
                 cabinet.inn, cabinet.fio, info)

        if not info:
            info = ('',) * 8
        else:
            assert len(info) == 8

        headers = ['inn', 'fio', 'parsed',  # status may be 0/1/2
                   'report', 'status', 'year', 'month', 'sent_date', 'sent_time', 'result_text',
                   'doc_state']
        row = [cabinet.inn, cabinet.fio, datetime.now(),
               info[0], info[1], info[2], info[3], info[4], info[5], info[6], info[7]]

        if inn in inn_rows:
            write_row_by_index_xls(filename, inn_rows[inn], row)
        else:
            append_xls(filename, headers, row)


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
    funcs = ['scan_keys', 'get_budget_status', 'get_receipts_status', 'get_payer_info',
             'send_outbox']
    try:
        func = choice.Menu(funcs).ask()
        globals()[func]()
    except Exception as e:
        log.exception(repr(e))
    input('DONE. press any key to close')
