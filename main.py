# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import os

import xlrd
import xlwt
import csv
import sqlite3

data_path = './CustomerData/'

amazon_top_negative_rating_book_path = data_path + 'Amazon TOP negative rating of Y&U series.xlsx'
oneplus_tv_order_id_book_path = data_path + 'OnePlus TV OrderID Reviews All Dump 25Aug2020.xlsx'
amazon_customer_details_path = data_path + 'InstallationReport/'
# amazon_customer_details_path = data_path + 'InstallationReportTest/'

empty_order_id = '000-0000000-0000000'

output_title = ['Subject', 'OrderId', 'Customer Name', 'Customer Address', 'Customer Phone', 'CustomerEmail']


def get_row_values_on_sheet(sheet, column):
    # print(sheet.name, column)
    column_result = []
    column_index = -1
    for row_index in range(sheet.nrows):
        if row_index == 0:
            column_index = sheet.row_values(row_index).index(column)
        else:
            if column_index != -1:
                column_result.append(sheet.row_values(row_index)[column_index])

    return column_result


def get_sheets_in_book(book):
    sheets = []
    for sheet_item in book.sheets():
        sheets.append(sheet_item)
    return sheets


def get_sheet0_in_book(book):
    return book.sheets()[0]


def open_book(filename):
    return xlrd.open_workbook(filename)


def get_subject_order_ids_result(subjects):
    order_id_result = []
    order_id_book = open_book(oneplus_tv_order_id_book_path)
    sheet = get_sheet0_in_book(order_id_book)
    review_titles = get_row_values_on_sheet(sheet, 'review_title')
    order_ids = get_row_values_on_sheet(sheet, 'Order Id')
    for subject in subjects:
        if subject in review_titles:
            index = review_titles.index(subject)
            order_id_result.append(order_ids[index])
        else:
            order_id_result.append(empty_order_id)
    return order_id_result


def get_amazon_customer_details_paths():
    file_paths = []
    details_file_names = os.listdir(amazon_customer_details_path)
    for file_name in details_file_names:
        file_paths.append(amazon_customer_details_path + file_name)
    return file_paths


def checkout_details_by_order_id_in_sheet(conn, subjects, order_ids):
    details = []

    cursor = conn.cursor()
    for (subject, order_id) in zip(subjects, order_ids):
        details_item = {'subject': subject, 'order_id': order_id}
        cursor.execute(
            """SELECT "Customer Name","Customer Address","Customer Phone",CustomerEmail FROM result where OrderId=?""",
            (order_id,))
        results = cursor.fetchall()
        if len(results) > 0:
            result = results[0]
            print('result=>>>>>', result)
            details_item['customer_name'] = result[0]
            details_item['customer_address'] = result[1]
            details_item['customer_phone'] = result[2]
            details_item['customer_email'] = result[3]
            details.append(details_item)
    cursor.close()
    return details


def merge_subjects_and_order_id(subjects, order_ids):
    details = []
    for (subject, order_id) in zip(subjects, order_ids):
        details.append({'subject': subject, 'order_id': order_id})
    return details


def checkout_details_by_order_ids_csv(subjects, order_ids):
    details = []
    details_file_paths = get_amazon_customer_details_paths()
    for file_path in details_file_paths:
        with open(file_path, 'r') as details_file:
            reader = csv.reader(details_file)
            temp_details = []
            sheet_order_ids = []
            for row in reader:
                sheet_order_ids.append(row[0])

            details.extend(temp_details)

    return details


def output_result(work_book, sheet_name, details):
    worksheet = work_book.add_sheet(sheet_name)

    row_index = 0
    column_index = 0
    for title in output_title:
        worksheet.col(column_index).width = 256 * (len(title) * 3 + 1)
        worksheet.write(row_index, column_index, title)
        column_index += 1
    for detail in details:
        print(detail)
        row_index += 1
        worksheet.write(row_index, 0, detail['subject'])
        worksheet.write(row_index, 1, detail['order_id'])
        worksheet.write(row_index, 2, detail['customer_name'])
        worksheet.write(row_index, 3, detail['customer_address'])
        worksheet.write(row_index, 4, detail['customer_phone'])
        worksheet.write(row_index, 5, detail['customer_email'])


def open_database():
    return sqlite3.connect('data.db')


def init_database(conn):
    cursor = conn.cursor()

    cursor.execute('DROP TABLE IF EXISTS result')
    cursor.execute('create table IF NOT EXISTS result (OrderId TEXT primary key'
                   ', Category TEXT'
                   ', Brand TEXT'
                   ', "Model Number" TEXT'
                   ', Title TEXT'
                   ', PostalCode TEXT'
                   ', City TEXT'
                   ', Quantity TEXT'
                   ', "Customer Name" TEXT'
                   ', "Customer Address" TEXT'
                   ', "Customer Phone" TEXT'
                   ', "CustomerEmail" TEXT'
                   ', "Estimated Delivery Date" TEXT'
                   ', "Service Request No." TEXT'
                   ', "Status" TEXT'
                   ')')
    cursor.close()
    conn.commit()


def get_installation_report_files(file_path, file_type='.csv'):
    name = []
    for root, dirs, files in os.walk(file_path):
        for i in files:
            if file_type in i:
                name.append(file_path + i)
    return name


def import_reporter_data(conn):
    cursor = conn.cursor()
    details_file_paths = get_installation_report_files(amazon_customer_details_path)

    for file_path in details_file_paths:
        if os.path.isfile(file_path):
            with open(file_path, 'r', encoding='UTF-8') as details_file:
                reader = csv.reader(details_file)
                row_index = 0
                for row in reader:
                    if row_index > 0:
                        data = ','.join('"%s"' % i for i in row)
                        placeholder = ','.join('?' for i in row)
                        stmt = 'INSERT INTO result VALUES (%s)' % placeholder
                        # print('===>', stmt)
                        try:
                            cursor.execute(stmt, row)
                            print('import data ===>', data)
                        except Exception:
                            print("import ", data, "error")

                    row_index += 1
        conn.commit()
    cursor.close()


def analise(conn):
    result_book = xlwt.Workbook(encoding='utf-8')

    if not os.path.exists(amazon_top_negative_rating_book_path):
        print("negative rating data not exist")
        return

    negative_rating_book = open_book(amazon_top_negative_rating_book_path)

    if not os.path.exists(oneplus_tv_order_id_book_path):
        print("OnePlus TV OrderID data not exist")
        return

    book_sheets = get_sheets_in_book(negative_rating_book)
    for sheet in book_sheets:
        sheet_name = sheet.name
        print(sheet_name)
        subjects = get_row_values_on_sheet(sheet, 'Subject')
        # print(subjects)
        order_ids = get_subject_order_ids_result(subjects)
        # print(order_ids)
        customer_details = checkout_details_by_order_id_in_sheet(conn, subjects, order_ids)
        print(customer_details)
        output_result(result_book, sheet.name, customer_details)
    result_book.save('output.xls')


def initEnv():
    if not os.path.exists(data_path):
        os.mkdir(data_path)
    if not os.path.exists(amazon_customer_details_path):
        os.mkdir(amazon_customer_details_path)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    initEnv()
    reporter_files = get_installation_report_files(amazon_customer_details_path)
    for file in reporter_files:
        print(file)
    conn = open_database()
    init_database(conn)
    import_reporter_data(conn)
    analise(conn)
    conn.close()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
