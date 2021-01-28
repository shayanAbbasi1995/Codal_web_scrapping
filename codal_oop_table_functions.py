from .codal_normal_functions import *


class cell:
    def __init__(self):
        self.rowspan = 0
        self.colspan = 0
        self.data = ''

    def add_rowspan(self, rowspan):
        self.rowspan = int(rowspan)

    def add_colspan(self, colspan):
        self.colspan = int(colspan)

    def add_data(self, data):
        self.data = data

    def print_span(self):
        print(self.rowspan, self.colspan)

    def print_data(self):
        print(self.data)

    def give_rowspan(self):
        return self.rowspan

    def give_colspan(self):
        return self.colspan

    def give_data(self):
        return self.data


class table:
    @staticmethod
    def find_colspan(data, line):
        try:
            data.add_colspan(line.get('colspan'))
        except:
            data.add_colspan(0)

    @staticmethod
    def find_rowspan(data, line):
        try:
            data.add_rowspan(line.get('rowspan'))
        except:
            data.add_rowspan(0)

    @staticmethod
    def class_attribute(line):
        try:
            class_element = line.get('class')
            if class_element == None:
                return False
            if isinstance(class_element, list):
                cul = ''
                for w in class_element:
                    cul += w
                class_element = cul
            if class_element.find('Hidden') != -1:
                return True
            else:
                return False
        except:
            return None

    @staticmethod
    def hidden_attribute(line):
        try:
            hidden_element = line.get('hidden')
            if hidden_element == None:
                return False
            else:
                return True
        except:
            return False

    @staticmethod
    def style_attribute(line):
        try:
            style_element = line.get('style')
            if style_element == None:
                return False
            if style_element.find('display:none') != -1:
                return True
            else:
                return False
        except:
            return False

    @staticmethod
    def input_tag(line):
        try:
            if line.find('span') != None:
                return False
            input_element = line.find('input')
            if input_element == None:
                return False
            type_element = input_element.get('type')
            if type_element == None:
                return False
            x = input_element['value']
            if type_element == 'text':
                return True
            else:
                return False
        except:
            return False

    @staticmethod
    def get_cols(row, row_type):
        satr = []
        if row_type == 'th':
            header = True
        else:
            header = False
        for col in row.find_all(row_type):
            one_cell = cell()
            table.find_colspan(one_cell, col)
            table.find_rowspan(one_cell, col)
            if table.class_attribute(col):
                continue
            if table.hidden_attribute(col):
                continue
            if table.style_attribute(col):
                continue
            if table.input_tag(col):
                input_element = col.find('input')
                one_cell.add_data(rep_char(input_element['value'], header))
            else:
                one_cell.add_data(rep_char(col.get_text(), header))
            satr.append(one_cell)
        return satr

    @classmethod
    def get_table(cls, table, hole_data):
        class_table = cls()
        for row in table.find_all("tr"):
            if cls.class_attribute(row):
                continue
            if len(row.find_all("th")) != 0:
                satr = class_table.get_cols(row, 'th')
            else:
                satr = class_table.get_cols(row, 'td')
            hole_data.append(satr)
        return hole_data


class codal_table:
    @staticmethod
    def first_info_table(table):
        if str(table).find('PeriodExtraDay') != -1:
            return True
        else:
            return False

    @staticmethod
    def body_table(soup):
        my_table = soup.find_all('table')
        if codal_table.first_info_table(my_table):
            my_table.pop(0)
        t_body = my_table[0].find("tbody")
        hole_data = []
        hole_data = table.get_table(t_body, hole_data)
        return hole_data

    @staticmethod
    def head_body_table(soup):
        my_table = soup.find('table')
        t_head = my_table.find("thead")
        hole_data = []
        hole_data = table.get_table(t_head, hole_data)
        t_body = my_table.find("tbody")
        hole_data = table.get_table(t_body, hole_data)
        return hole_data

    @staticmethod
    def two_body_table(soup):
        t_body = soup.find_all("tbody")
        if codal_table.first_info_table(t_body):
            t_body.pop(0)
        hole_data = []
        hole_data = table.get_table(t_body[0], hole_data)
        hole_data = table.get_table(t_body[1], hole_data)
        return hole_data

    @staticmethod
    def two_table(soup):
        if len(soup.find_all('table')) != 3:
            raise Exception("not a two_table")
        main_table = soup.find_all('table')
        if codal_table.first_info_table(main_table):
            main_table.pop(0)
        t_body = main_table[1].find("tbody")
        hole_data = []
        hole_data = table.get_table(t_body, hole_data)
        t_body = main_table[2].find("tbody")
        hole_data = table.get_table(t_body, hole_data)
        return hole_data

    @classmethod
    def final_table(cls, soup):
        try:
            x = cls.head_body_table(soup)
            print('head_body_table')
            return x
        except:
            try:
                x = cls.two_table(soup)
                print('two_table')
                return x
            except:
                try:
                    x = cls.two_body_table(soup)
                    print('two_body_table')
                    return x
                except:
                    x = cls.body_table(soup)
                    print('body_table')
                    return x
class make_file :
    @staticmethod
    def make_empty_cells(data):
        i = 0
        empty_cell = cell()
        while i < len(data):
            j = 0
            while j < len(data[i]):
                for k in range(1, data[i][j].give_rowspan()):
                    data[i + k].insert(j, empty_cell)
                for k in range(data[i][j].give_colspan() - 1):
                    data[i].insert(j + 1, empty_cell)
                j += 1
            i += 1
        return data

    @staticmethod
    def make_excel(name, stock_id, data, attribute):
        wb = Workbook()
        wb.save('codal'+ '\\' + stock_id + '\\' + name + '.xlsx')

        x = []
        x.append(name)
        x += attribute
        ws = wb.active
        ws.title = "Page 1"
        ws.append(x)
        data = make_file.make_empty_cells(data)
        for i in range(len(data)):
            temp = []
            for j in range(len(data[i])):
                try:
                    temp.append(str_to_int_or_float(data[i][j].give_data()))
                except:
                    temp.append(data[i][j].give_data())
            ws.append(temp)
        wb.save('codal'+ '\\' + stock_id + '\\' + name + '.xlsx')

    @staticmethod
    def make_folder(id):
        try:
            os.mkdir('codal')
        except:
            pass
        try:
            os.mkdir('codal' + '\\' +str(id))
        except:
            pass

    @staticmethod
    def make_sub_file(stock_id, company, sub_id, company_state_code):
        try:
            wb = openpyxl.load_workbook('codal'+ '\\' + stock_id + '\\' + stock_id + ".xlsx")
            ws = wb.active
        except:
            wb = Workbook()
            wb.save('codal'+ '\\' + stock_id + '\\' + stock_id + ".xlsx")
            ws = wb.active
            ws.title = "Page 1"
            ws.cell(row=1, column=1).value = 'Stock id'
            ws.cell(row=1, column=2).value = 'Main company name'
            ws.cell(row=1, column=3).value = 'Company status'
            ws.cell(row=2, column=1).value = stock_id
            if sub_id is None:
                ws.cell(row=2, column=2).value = str(company)
                ws.cell(row=2, column=3).value = company_state_code
            ws.cell(row=3, column=1).value = 'Sub id'
            ws.cell(row=3, column=2).value = 'Sub company name'
            ws.cell(row=3, column=3).value = 'Company status'
            for i in range(1, 10):
                ws.cell(row=i + 3, column=1).value = '0' + str(i)
            for j in range(10, 100):
                ws.cell(row=j + 3, column=1).value = str(j)

        if sub_id == None and ws.cell(row=2, column=2).value == None:
            ws.cell(row=2, column=2).value = str(company)
            ws.cell(row=2, column=3).value = company_state_code
        if sub_id != None:
            for i in range(1, 100):
                if int(sub_id) == i and ws.cell(row=i + 3, column=2).value == None:
                    ws.cell(row=i + 3, column=2).value = str(company)
                    ws.cell(row=i + 3, column=3).value = company_state_code
        wb.save('codal'+ '\\' + stock_id + '\\' + stock_id + ".xlsx")

    @staticmethod
    def make_error_file(stock_id, stock_type, link):
        try:
            wb = openpyxl.load_workbook('codal'+ '\\' + stock_id + '\\' + stock_id + '_errors' + ".xlsx")
            ws = wb.active
        except:
            wb = Workbook()
            wb.save('codal'+ '\\' + stock_id + '\\' + stock_id + '_errors' + ".xlsx")
            ws = wb.active
            ws.title = "Page 1"
        x = []
        print('Found unreadable table , link :', link)
        x.append(stock_type)
        x.append(link)
        ws.append(x)
        wb.save('codal'+ '\\' + stock_id + '\\' + stock_id + '_errors' + ".xlsx")