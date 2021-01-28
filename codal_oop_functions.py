from .codal_oop_table_functions import *


class stock_codal:
    second = 0.001
    path = "chromedriver.exe"
    url = "https://codal.ir/"
    codal_loading = 'col-12 loading ng-scope'
    stock_file_names = 'stock_names.xlsx'
    id_column = 0
    name_column = 1
    first_page_url = 'https://codal.ir/ReportList.aspx?search&Symbol=findme&LetterType=-1&Isic=722008&AuditorRef=-1&PageNumber=1&Audited&NotAudited&IsNotAudited=false&Childs&Mains&Publisher=false&CompanyState=0&Category=-1&CompanyType=1&Consolidatable&NotConsolidatable'

    company_id = 'ctl00_txbCompanyName'
    capital_id = 'ctl00_lblListedCapital'
    symbol_id = 'ctl00_txbSymbol'
    unauthorized_capital_id = 'ctl00_txbUnauthorizedCapital'
    ISIC_id = 'ctl00_lblISIC'
    period_id = 'ctl00_lblPeriod'
    period_end_id = 'ctl00_lblPeriodEndToDate'
    date_id = 'ctl00_lblYearEndToDate'
    company_state_id = 'ctl00_lblCompanyState'

    def __init__(self, my_name, my_id):
        self.name = my_name
        self.id = my_id
        self.status = True
        self.all_activities = []
        self.num_activities = 0
        self.all_fiscals = []
        self.num_fiscals = 0

    def get_activity(self, my_activity):
        self.all_activities.append(my_activity)
        self.num_activities += 1

    def get_fiscal(self, my_fiscal):
        self.all_fiscals.append(my_fiscal)
        self.num_fiscals += 1

    def remove_activity(self, index):
        print('removeeeeeeeeeeeeeeeeeeeeeee', self.all_activities.pop(index))
        self.num_activities -= 1

    def remove_fiscal(self, index):
        print('removeeeeeeeeeeeeeeeeeeeeeee', self.all_fiscals.pop(index))
        self.num_fiscals -= 1

    def print_name(self):
        print('** name of stock is', self.name)

    def print_all_activities(self):
        print('** list of activities for stock', self.name, ':')
        for i in range(len(self.all_activities)):
            print('    ', i + 1, self.all_activities[i])

    def print_all_fiscals(self):
        print('** list of fiscals for stock', self.name, ':', )
        for i in range(len(self.all_fiscals)):
            print('    ', i + 1, self.all_fiscals[i])

    def print_num_fiscals(self):
        print('Stock ', self.name, ' have ', self.num_fiscals, 'fiscals')

    def print_num_activities(self):
        print('Stock ', self.name, ' have ', self.num_activities, 'activities')

    def open_first_page(self, browser):
        browser.get(self.first_page_url.replace('findme', self.name))
        time.sleep(15)

    @classmethod
    def get_stock_names(cls, path, id_column, name_column, start, end):
        wb = openpyxl.load_workbook(path)
        sheet = wb.active
        names = list(sheet.columns)[name_column]
        ids = list(sheet.columns)[id_column]
        x = []
        for i in range(start, end + 1):
            if names[i - 1].value != None:
                x.append(cls(str(names[i - 1].value), str(ids[i - 1].value)))
        return x

    @staticmethod
    def find_all_li(browser):
        page = browser.page_source
        soup = BeautifulSoup(page, 'lxml')
        my_nav = soup.find_all('nav')[1]
        my_ul = my_nav.find('ul')
        li = my_ul.find_all("li")
        return li

    @staticmethod
    def find_num_pages(li):
        number_of_pages = li[len(li) - 4].find('a')
        number_of_pages = number_of_pages.get_text()
        number_of_pages = change_numbers(number_of_pages)
        number_of_pages = int(number_of_pages)
        return number_of_pages

    @staticmethod
    def find_date(line):
        match = re.search(r'\d{4}-\d{2}-\d{2}', line)
        if match is None:
            match = re.search(r'\d{2}-\d{2}-\d{4}', line)
        if match is None:
            match = re.search(r'\d{2}-\d{2}-\d{2}', line)
        if match is None:
            print("ERROR___there is sth wrong with data:", line)
        return match

    @staticmethod
    def find_period(line):
        if line.find('1') != -1:
            return 1
        elif line.find('2') != -1:
            return 2
        elif line.find('3') != -1:
            return 3
        elif line.find('6') != -1:
            return 6
        elif line.find('9') != -1:
            return 9
        elif line.find('12') != -1:
            return 12
        else:
            return None

    @staticmethod
    def find_link(line):
        line_a = line.find("a")
        link = line_a['href']
        link = 'https://codal.ir' + link
        return link

    @staticmethod
    def find_sub(line, audit, correction):
        if correction:
            y = line.find('(اصلاحیه)')
            x = line[y + 9:]
        else:
            if audit:
                y = line.find('(حسابرسی شده)')
            else:
                y = line.find('(حسابرسی نشده)')
            x = line[y + 14:]
        if x == '':
            return "main"
        else:
            return x

    @staticmethod
    def check_state(state):
        if state.find('فرابورس') != -1:
            return 'Stock morket'
        elif state.find('بورس') != -1:
            return 'OTC market'
        elif state.find('نشده'):
            return 'Not accepted'
        else:
            return 'Other'

    @staticmethod
    def find_month(date):
        if date.find('/01/') != -1:
            return '01'
        elif date.find('/02/') != -1:
            return '02'
        elif date.find('/03/') != -1:
            return '03'
        elif date.find('/04/') != -1:
            return '04'
        elif date.find('/05/') != -1:
            return '05'
        elif date.find('/06/') != -1:
            return '06'
        elif date.find('/07/') != -1:
            return '07'
        elif date.find('/08/') != -1:
            return '08'
        elif date.find('/09/') != -1:
            return '09'
        elif date.find('/10/') != -1:
            return '10'
        elif date.find('/11/') != -1:
            return '11'
        elif date.find('/12/') != -1:
            return '12'

    @staticmethod
    def open_browser(path):
        option = Options()
        option.add_argument("--disable-infobars")
        option.add_argument("start-maximized")
        option.add_argument("--disable-extensions")
        option.add_experimental_option("prefs", {"profile.default_content_setting_values.notifications": 1})
        return webdriver.Chrome(options=option, executable_path=path)

    def open_codal(self, browser):
        while True:
            try:
                browser.get(self.url)
                print("SUCCESS___I could open the main page.")
                time.sleep(1)
                search = browser.find_element_by_id('aSearch')
                search.click()
                print("SUCCESS___I could click on search.")
                time.sleep(1)
                search_space = browser.find_element_by_xpath('//*[@id="collapse-search-1"]/div[2]/div[1]/div/div/a')
                search_space.click()
                time.sleep(1)
                search_box = browser.find_element_by_id('txtSymbol')
                search_box.send_keys(self.name)
                break
            except:
                print("ERROR___I could NOT open the main page.")
                time.sleep(1)

    def get_search_result(self, browser):
        while True:
            try:
                load_element = browser.find_element_by_class_name(self.codal_loading)
                time.sleep(stock_codal.second)
            except:
                search_result = None
                try:
                    search_result = browser.find_element_by_xpath(
                        '//*[@id="ui-select-choices-row-0-0"]/div/div[1]/span')
                except:
                    try:
                        search_result = browser.find_element_by_xpath('//*[@id="ui-select-choices-row-0-0"]/div/div[1]')
                    except:
                        try:
                            search_result = browser.find_element_by_xpath(
                                '//*[@id="ui-select-choices-row-0-0"]/div/div[2]')
                        except:
                            print('cant search stock', self.name)
                return search_result

    def is_it_bad_stock(self, browser, search_result):
        times = 0
        while True:
            try:
                load_element = browser.find_element_by_class_name(self.codal_loading)
                time.sleep(self.second)
            except:
                try:
                    search_result.click()
                    return False
                except:
                    if times == 5:
                        self.status = False
                        return True
                        print(self.name, 'is a bad stock')
                    else:
                        time.sleep(1)
                        times += 1

    def check_get_li(self, browser):
        while True:
            try:
                load_element = browser.find_element_by_class_name(self.codal_loading)
                time.sleep(self.second)
            except:
                try:
                    return self.find_all_li(browser)
                except:
                    time.sleep(self.second)

    def next_page_address(self, browser, stock_url, counting):
        y = stock_url.find('PageNumber')
        stock_url = stock_url[:y + 11] + str(counting + 1) + stock_url[y + 12:]
        while True:
            try:
                load_element = browser.find_element_by_class_name(self.codal_loading)
                time.sleep(self.second)
            except:
                return stock_url

    def accessibility_func(self, browser):
        times = 0
        number_of_pages = 100000
        while True:
            try:
                load_element = browser.find_element_by_class_name(self.codal_loading)
                time.sleep(self.second)
            except:
                try:
                    li = self.find_all_li(browser)
                    number_of_pages = self.find_num_pages(li)
                except:
                    if times == 5:
                        print("Page", self.name, 'Not accessable')
                        self.status = False
                        return False
                    else:
                        time.sleep(1)
                        times += 1
                if number_of_pages <= 5000:
                    print("namad", self.name, "has", number_of_pages, "pages")
                    return True

    def open_report_page(self, url, browser):
        while True:
            try:
                browser.get(url)
                print("success___I could open the link page")
                break
            except:
                print("ERROR___I could NOT open the link page :", url)
                time.sleep(self.second)
        while True:
            try:
                main_page = browser.page_source
                main_soup = BeautifulSoup(main_page, 'lxml')
                main_menu = main_soup.find_all('option')
                print("success___I could make a soup")
                return main_menu, browser
            except:
                print("ERROR___I could not make a soup")
                time.sleep(self.second)

    def click_menu(self, browser, i):
        try:
            menu = browser.find_element_by_xpath('/html/body/form/div[4]/div[3]/select/option[' + str(i) + ']')
            menu.click()
        except:
            try:
                menu = browser.find_element_by_xpath('/html/body/form/div[4]/div[3]/select/option')
                menu.click()
            except:
                try:
                    menu = browser.find_element_by_xpath('/html/body/form/div[3]/div[3]/select/option[' + str(i) + ']')
                    menu.click()
                except:
                    try:
                        menu = browser.find_element_by_xpath('/html/body/form/div[3]/div[3]/select/option')
                        menu.click()
                    except:
                        pass


class activity(stock_codal):
    def __init__(self):
        self.company = None
        self.symbol = None
        self.capital = None
        self.unauthorized_capital = None
        self.period_end = None
        self.date = None
        self.company_state = None
        self.link = None

    def print_activity(self):
        print(self.symbol, self.company, self.period_end \
              , self.capital, self.unauthorized_capital, self.date, self.company_state, self.link)

    def print_list_of_activities(self, list_act):
        for i in list_act:
            print(i.date, i.link)

    def print_link(self):
        print(self.my_link)

    def add_link(self, link):
        self.link = link

    def add_company(self, soup):
        company_name = soup.find(id=self.company_id)
        if company_name is None:
            return
        company_name = company_name.get_text()
        self.company = str(company_name)

    def add_symbol(self, soup):
        symbol_name = soup.find(id=self.symbol_id)
        if symbol_name is None:
            return
        symbol_name = symbol_name.get_text()
        self.symbol = str(symbol_name)

    def add_capital(self, soup):
        capital_name = soup.find(id=self.capital_id)
        if capital_name is None:
            return
        capital_name = capital_name.get_text()
        capital_name = capital_name.replace(',', '')
        self.capital = int(capital_name)

    def add_unauthorized_capital(self, soup):
        unauthorized_capital_name = soup.find(id=self.unauthorized_capital_id)
        if unauthorized_capital_name is None:
            return
        unauthorized_capital_name = unauthorized_capital_name.get_text()
        unauthorized_capital_name = unauthorized_capital_name.replace(',', '')
        self.unauthorized_capital = str(unauthorized_capital_name)

    def add_period_end(self, soup):
        period_end_name = soup.find(id=self.period_end_id)
        if period_end_name is None:
            return
        period_end_name = period_end_name.get_text()
        period_end_name = self.find_month(period_end_name)
        self.period_end = str(period_end_name)

    def add_date(self, soup):
        date_name = soup.find(id=self.date_id)
        if date_name is None:
            return
        date_name = date_name.get_text()
        y = date_name.find('13')
        date_name = date_name[y:y + 4]
        self.date = str(date_name)

    def add_company_state(self, soup):
        company_state_name = soup.find(id=self.company_state_id)
        if company_state_name is None:
            return
        company_state_name = company_state_name.get_text()
        company_state_name = self.check_state(company_state_name)
        self.company_state = str(company_state_name)

    def make_name(self):
        name = str(self.date) + '-' + str(self.period_end)
        return name

    def save_attribute(self):
        att = [self.company, self.symbol, self.capital, self.unauthorized_capital, self.company_state]
        return att

    def sub_id(self):
        x = re.search(r'\d+', self.symbol)
        if x is None:
            return '00'
        elif int(x.group()) < 10:
            return '0' + x.group()
        elif int(x.group()) >= 10:
            return x.group()

    def code_company_state(self):
        if self.company_state == 'Stock morket':
            return '01'
        elif self.company_state == 'OTC market':
            return '02'
        elif self.company_state == 'Other':
            return '03'
        else:
            return 'N/A'

    def check_existence(self):
        if self.company is None:
            return False
        if self.symbol is None:
            return False
        if self.capital is None:
            return False
        return True

    def activity_analysis(self, soup):
        self.add_company(soup)
        self.add_symbol(soup)
        self.add_capital(soup)
        self.add_unauthorized_capital(soup)
        self.add_period_end(soup)
        self.add_date(soup)
        self.add_company_state(soup)

    @staticmethod
    def page_not_found(browser):
        if browser.current_url.find('ErrorMsg') == -1:
            return False
        else:
            return True

    def get_data_make_table(self, browser, stock_id, code):
        try:
            page = browser.page_source
            soup = BeautifulSoup(page, 'lxml')
            if self.page_not_found(browser):
                return
            hole_data = codal_table.final_table(soup)
            self.save_attribute().insert(0, str(browser.current_url))
            make_file.make_excel(stock_id + '-' + self.sub_id() + '-' + code + '0-' + self.make_name(), stock_id, hole_data,
                       self.save_attribute())
        except:
            make_file.make_error_file(stock_id, code, self.link)

    def give_activity_link_get_table(self, browser, stock_id):
        main_menu, browser = self.open_report_page(self.link, browser)
        print('activity report for link :', self.link)
        for i in range(1, len(main_menu) + 1):
            self.click_menu(browser, i)
            if main_menu[i - 1].get_text().find('گزارش فعالیت ماهانه') != -1:
                self.get_data_make_table(browser, stock_id, '01-')

    def symbol_analysis(self, stock_name, stock_id):
        if self.symbol is None:
            return
        print(stock_name, stock_id, self.symbol, self.link)
        print(self.link)
        sub_id = re.search(r'\d+', self.symbol)
        if sub_id is None:
            make_file.make_sub_file(stock_id, self.company, None, self.code_company_state())
        else:
            make_file.make_sub_file(stock_id, self.company, sub_id.group(), self.code_company_state())


class fiscal(stock_codal):
    def __init__(self):
        self.correction = None
        self.audit = None
        self.company = None
        self.symbol = None
        self.capital = None
        self.unauthorized_capital = None
        self.period = None
        self.term = None
        self.period_end = None
        self.date = None
        self.company_state = None
        self.link = None

    def print_fiscal(self):
        print(self.symbol, self.audit, self.correction, self.company, self.period_end, self.term, self.period \
              , self.capital, self.unauthorized_capital, self.date, self.company_state, self.link)

    def give_link(self, link):
        self.my_link = link

    def print_link(self):
        print(self.my_link)

    def add_correction(self, correction):
        self.correction = correction

    def add_link(self, link):
        self.link = link

    def add_term(self, term):
        self.term = term

    def add_audit(self, audit):
        self.audit = audit

    def add_company(self, soup):
        company_name = soup.find(id=self.company_id)
        if company_name is None:
            return
        company_name = company_name.get_text()
        self.company = str(company_name)

    def add_symbol(self, soup):
        symbol_name = soup.find(id=self.symbol_id)
        if symbol_name is None:
            return
        symbol_name = symbol_name.get_text()
        self.symbol = str(symbol_name)

    def add_capital(self, soup):
        capital_name = soup.find(id=self.capital_id)
        if capital_name is None:
            return
        capital_name = capital_name.get_text()
        capital_name = capital_name.replace(',', '')
        self.capital = int(capital_name)

    def add_unauthorized_capital(self, soup):
        unauthorized_capital_name = soup.find(id=self.unauthorized_capital_id)
        if unauthorized_capital_name is None:
            return
        unauthorized_capital_name = unauthorized_capital_name.get_text()
        unauthorized_capital_name = unauthorized_capital_name.replace(',', '')
        self.unauthorized_capital = str(unauthorized_capital_name)

    def add_period(self, soup):
        period_name = soup.find(id=self.period_id)
        if period_name is None:
            return
        period_name = period_name.get_text()
        period_name = self.find_period(period_name)
        self.period = int(period_name)

    def add_period_end(self, soup):
        period_end_name = soup.find(id=self.period_end_id)
        if period_end_name is None:
            return
        period_end_name = period_end_name.get_text()
        period_end_name = self.find_month(period_end_name)
        self.period_end = str(period_end_name)

    def add_date(self, soup):
        date_name = soup.find(id=self.date_id)
        if date_name is None:
            return
        date_name = date_name.get_text()
        y = date_name.find('13')
        date_name = date_name[y:y + 4]
        self.date = str(date_name)

    def add_company_state(self, soup):
        company_state_name = soup.find(id=self.company_state_id)
        if company_state_name is None:
            return
        company_state_name = company_state_name.get_text()
        company_state_name = self.check_state(company_state_name)
        self.company_state = str(company_state_name)

    def make_name(self):
        name = ''
        if self.audit:
            name += '1-'
        else:
            name += '0-'
        if self.correction:
            name += '1-'
        else:
            name += '0-'
        name += str(self.date) + '-' + str(self.period_end) + '-' + str(self.period)
        return name

    def save_attribute(self):
        att = [self.company, self.symbol, self.capital, self.unauthorized_capital, self.company_state]
        return att

    def code_company_state(self):
        if self.company_state == 'Stock morket':
            return '01'
        elif self.company_state == 'OTC market':
            return '02'
        elif self.company_state == 'Other':
            return '03'
        else:
            return 'N/A'

    def sub_id(self):
        x = re.search(r'\d+', self.symbol)
        if x is None:
            return '00'
        elif int(x.group()) < 10:
            return '0' + x.group()
        elif int(x.group()) >= 10:
            return x.group()

    def check_existence(self):
        if self.company is None:
            return False
        if self.symbol is None:
            return False
        if self.capital is None:
            return False
        return True

    def fiscal_analysis_1(self, line_span, line_td, audit):
        if audit:
            self.add_audit(True)
        else:
            self.add_audit(False)
        if line_span.find('(اصلاحیه)') != -1:
            self.add_correction(True)
        else:
            self.add_correction(False)
        self.add_link(self.find_link(line_td))

    def fiscal_analysis_2(self, soup):
        self.add_company(soup)
        self.add_symbol(soup)
        self.add_capital(soup)
        self.add_unauthorized_capital(soup)
        self.add_period(soup)
        self.add_period_end(soup)
        self.add_date(soup)
        self.add_company_state(soup)

    @staticmethod
    def page_not_found(browser):
        if browser.current_url.find('ErrorMsg') == -1:
            return False
        else:
            return True

    def get_data_make_table(self, browser, stock_id, code, main_menu):
        try:
            page = browser.page_source
            soup = BeautifulSoup(page, 'lxml')
            if self.page_not_found(browser):
                return
            hole_data = codal_table.final_table(soup)
            self.save_attribute().insert(0, browser.current_url)
            if main_menu != -1:
                make_file.make_excel(stock_id + '-' + self.sub_id() + '-' + code + '1-' + self.make_name(), stock_id, hole_data,
                           self.save_attribute())
            else:
                make_file.make_excel(stock_id + '-' + self.sub_id() + '-' + code + '0-' + self.make_name(), stock_id, hole_data,
                           self.save_attribute())
        except:
            make_file.make_error_file(stock_id, code, self.link)

    def give_fiscal_link_get_table(self, browser, stock_id):
        main_menu, browser = self.open_report_page(self.link, browser)
        print('fiscal report for link :', self.link)
        for i in range(1, len(main_menu) + 1):
            self.click_menu(browser, i)
            talfig = main_menu[i - 1].get_text().find('تلفیقی')
            if main_menu[i - 1].get_text().find('ترازنامه') != -1:
                self.get_data_make_table(browser, stock_id, '02-', talfig)
            if main_menu[i - 1].get_text().find('صورت سود و زیان') != -1:
                self.get_data_make_table(browser, stock_id, '03-', talfig)
            if main_menu[i - 1].get_text().find('جریان وجوه نقد') != -1:
                self.get_data_make_table(browser, stock_id, '04-', talfig)

    def symbol_analysis(self, stock_name, stock_id):
        if self.symbol is None:
            return
        sub_id = re.search(r'\d+', self.symbol)
        if sub_id is None:
            make_file.make_sub_file(stock_id, self.company, None, self.code_company_state())
        else:
            make_file.make_sub_file(stock_id, self.company, sub_id.group(), self.code_company_state())