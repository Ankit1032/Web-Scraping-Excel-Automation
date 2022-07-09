# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import chromedriver_autoinstaller
import time
from bs4 import BeautifulSoup
import requests
import csv
import threading

def enter_website_product_id(shpid):
    website_link = '(website link....)'

    #making extracting headless
    options = Options()

    # this parameter tells Chrome that
    # it should be run without UI (Headless)
    options.headless = True

    # initializing webdriver for Chrome with our options
    browser = webdriver.Chrome(options=options)

    #browser = webdriver.Chrome()
    browser.get(website_link) #enter into website website
    browser.maximize_window() #Maximize the window
    time.sleep(2) #Pausing the code to let the page load
    shp_obj = browser.find_element('id', 'transactions-view-shp_objId') #the field id of product_id
    shp_obj.clear() #clearing the product_id field
    shp_obj.send_keys(shpid) #filling the product_id field with value
    go_button = browser.find_element('id', 'transactions-view-shp_0') #the id of product_id's GO button
    go_button.click() #clicking the 'GO' button
    time.sleep(4) #Pausing the code to let the page load properly


    # Beautiful Soup
    # providing the page_source to soup | lxml -> parser
    soup = BeautifulSoup(browser.page_source, 'lxml')
    # finding relevant ids in page souce to extract product_id_id from the page
    website_product_id = soup.find('div', class_='alert alert-info').find('strong').text
    # print("product_id ID: ", website_product_id)
    # print('-------------------------------------------------------------------------')

    website_table_body = soup.find('table', class_='table table-striped dataTable no-footer').find('tbody')
    website_rows = website_table_body.find_all('tr', {'role': 'row'})
    # print(website_rows[0].find_all('td')[2].text)#transaction no
    # print(website_rows[0].find_all('td')[7].text)#transaction_time update(utc)

    transaction_dict = {}
    for index in range(0, len(website_rows)):
        transaction = website_rows[index].find_all('td')[2].text.strip()
        transaction_time_update = website_rows[index].find_all('td')[7].text.strip()

        #below helps in extract the transaction_time update of only the first occurance of the transaction
        if transaction not in transaction_dict:
            transaction_dict[transaction] = transaction_time_update

    return transaction_dict


def transaction_workbook(sheet_name):
    #load workbook
    ev_workbook = load_workbook(sheet_name, data_only=True)
    sheet_obj = ev_workbook.active # Get workbook active sheet object
    print("X-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x")

    # Ascii of A = 65, B= 66 ... Z = 60
    sheet_transaction_list = [] #will be present in column headers staring from 'B'
    col_head_ascii = 65
    for i in range(1,sheet_obj.max_column):
        col = chr(col_head_ascii+i)
        col_head_value = sheet_obj[col+'1'].value
        sheet_transaction_list.append([col,col_head_value])

    print(sheet_transaction_list)
    row_sheet = 2  # used to fill row wise dates of every product_ids

    for i in range(2,sheet_obj.max_row+1):
        cell_index = 'A'+str(i)
        shp_id_value = sheet_obj[cell_index].value

        # call function to extract all transaction of that product_id
        website_dict = enter_website_product_id(shp_id_value)

        for j in range(0,len(sheet_transaction_list)):
            col_alpha = sheet_transaction_list[j][0] #B,C,D..
            transaction_val = str(sheet_transaction_list[j][1]) #6,24-1,361....

            sheet_cell_index = col_alpha+str(row_sheet)
            if transaction_val in website_dict:

                sheet_obj[sheet_cell_index].value = website_dict[transaction_val]
            else:
                sheet_obj[sheet_cell_index].value = 'transaction_Not_Available'
        print(row_sheet-1,". ",shp_id_value," data extracted")

        row_sheet = row_sheet + 1


    # iterate through excel and display data
    # for i in range(1, sheet_obj.max_row + 1):
    #     print("\n")
    #     print("Row ", i, " data :")
    #
    #     for j in range(1, sheet_obj.max_column + 1):
    #         cell_obj = sheet_obj.cell(row=i, column=j)
    #         print(cell_obj.value, end=" ")

    ev_workbook.save('Final_transactions_Date.xlsx')



        # print(shp_id_value)
        # for row in range(0,len(website_transactions_mat)):
        #     print(website_transactions_mat[row])
        # #loop through all transactions



if __name__ == '__main__':
    # chromedriver_autoinstaller.install()  # Check if the current version of chromedriver exists
    # # and if it doesn't exist, download it automatically,
    # # then add chromedriver to path
    #
    # driver = webdriver.Chrome()
    # driver.get("http://www.python.org")
    # assert "Python" in driver.title

    #main code
    sheet_name = "extract_productid_transaction_time_update_of_transactions.xlsx"


    #load workbook
    transaction_workbook(sheet_name)
