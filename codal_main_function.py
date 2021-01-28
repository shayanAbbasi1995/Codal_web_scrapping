import datetime
from .codal_oop_functions import *


begin_time = datetime.datetime.now()


def codal_search_for_links(start, end):
    names = stock_codal.get_stock_names(stock_codal.stock_file_names, stock_codal.id_column, stock_codal.name_column,
                                        start, end)
    browser = stock_codal.open_browser(stock_codal.path)
    bad_stock_names = []

    for k in names:
        make_file.make_folder(k.id)

    for stock in names:
        print(stock.id, stock.name)
        stock.open_first_page(browser)
        li = stock.find_all_li(browser)
        try:
            number_of_pages = stock.find_num_pages(li)
        except:
            print('No data on this page')
            continue
        stock_url = browser.current_url

        for counting in range(1, number_of_pages + 1):
            print('SUCCESS___I could open page', counting, '.')
            page = browser.page_source
            soup = BeautifulSoup(page, 'lxml')
            my_table = soup.find('table')
            t_body = my_table.find("tbody")

            for line_tr in t_body.find_all("tr"):
                fiscal_data = fiscal()
                activity_data = activity()
                line_td = line_tr.find_all("td")[3]
                line_span = line_td.find('span')
                line_span = line_span.get_text()
                line_span = line_span.replace('\n', '')

                if line_span.find('(حسابرسی نشده)') != -1 and \
                        line_span.find('پیش بینی') == -1 and line_span.find('(به پیوست)') == -1:
                    fiscal_data.fiscal_analysis_1(line_span, line_td, False)
                    stock.get_fiscal(fiscal_data)

                elif line_span.find('(حسابرسی شده)') != -1 and \
                        line_span.find('پیش بینی') == -1 and line_span.find('(به پیوست)') == -1:
                    fiscal_data.fiscal_analysis_1(line_span, line_td, True)
                    stock.get_fiscal(fiscal_data)

                elif line_span.find('گزارش فعالیت ماهانه') != -1:
                    activity_data.add_link(activity_data.find_link(line_td))
                    stock.get_activity(activity_data)

            print(browser.current_url)
            browser.get(stock.next_page_address(browser, stock_url, counting))
            time.sleep(2)
            print("SUCCESS___End extracting of page number " + str(counting) + " for stock " + stock.name)
            counting += 1
        print("End finding links of stock", stock.name)
    for stock in names:
        if not stock.status:
            continue
        print('Extraxting tag informations for stock ', stock.name)
        stock.print_num_activities()
        j = 0
        while j < stock.num_activities:
            browser.get(stock.all_activities[j].link)
            time.sleep(2)
            page = browser.page_source
            soup = BeautifulSoup(page, 'lxml')
            stock.all_activities[j].activity_analysis(soup)
            if not stock.all_activities[j].check_existence():
                stock.remove_activity(j)
            else:
                stock.all_activities[j].give_activity_link_get_table(browser, stock.id)
            j += 1
        u = 0
        stock.print_num_fiscals()
        while u < stock.num_fiscals:
            browser.get(stock.all_fiscals[u].link)
            time.sleep(2)
            page = browser.page_source
            soup = BeautifulSoup(page, 'lxml')
            stock.all_fiscals[u].fiscal_analysis_2(soup)
            if not stock.all_fiscals[u].check_existence():
                stock.remove_fiscal(u)
            else:
                stock.all_fiscals[u].give_fiscal_link_get_table(browser, stock.id)
            u += 1
    for stock in names:
        print('Extraxting tables for stock ', stock.name)
        for j in range(stock.num_activities):
            stock.all_activities[j].symbol_analysis(stock.name, stock.id)
            print('activity number ', j + 1, ':')
            stock.all_activities[j].print_activity()
        for u in range(stock.num_fiscals):
            stock.all_fiscals[u].symbol_analysis(stock.name, stock.id)
            print('fiscal number ', u + 1, ':')
            stock.all_fiscals[u].print_fiscal()
    browser.quit()
    print(datetime.datetime.now() - begin_time)
    print('The END')

