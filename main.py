import time
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
import pandas as pd

password = 'password'
timetosleep = 0.5
save_path = r"C:\Users\Administrator\Downloads"


#   login into accounting program 1C
def login1C():
    app = Application(backend='uia').start('Path to 1C.exe')
    win = app.Dialog
    win.Button6.click()
    win.Button6.set_focus()
    win.Edit.type_keys(password)
    win.OK.click()
    time.sleep(timetosleep)


#   i need to be in "making bills". This function do it
def getNeedSpace(left, func):   #  left = 'Продажи' func = 'Счета на оплату покупателям'
    app = Application(backend='uia').connect(title='Бухгалтерия для Украины, редакция 2.0. / ')  # get the app to work with
    win = app['Бухгалтерия для Украины, редакция 2.0. / Dialog']

    sales = win.descendants(title=left)[0].select() # selecting "sales"
    sales.click_input() # click on it
    time.sleep(timetosleep)
    send_keys('^f '+func, with_spaces=True) # typing name of needed function
    time.sleep(timetosleep)
    send_keys('{ENTER}{ENTER}')
    return win # return new working window

def operateSell():

    win = getNeedSpace('Продажи', 'Счета на оплату покупателям') # get new working window
    time.sleep(timetosleep)
    win.descendants(title='Создать')[0].click_input() # click to make new bill

    app = Application(backend='uia').connect(title='Бухгалтерия для Украины, редакция 2.0. ')
    win = app['Бухгалтерия для Украины, редакция 2.0. / Dialog']

    df = pd.read_excel(r'Path to excel file ') # get the excel to parse info

    size = len(df) # get total number of rows in excel file
    size = size - 2

    companies = df['Unnamed: 4'][size]   # get the least names of companies
                                         # names of companies stay in format name1 - name2
                                         # so i need to get this names in 2 variable
    comp1 = '' # name1
    comp2 = '' # name2
    f = 0
    for i in range(0, len(companies)): # make a loop
        if companies[i] == '-': # if i get this symbol, so name2 starts
            f = 1
        else:
            if f == 0:
                comp1 += companies[i]
            else:
                if companies[i] != ' ':
                    comp2 += companies[i]

    weight = str(df['Unnamed: 10'][size]) # get weight
    price = str(df['Unnamed: 11'][size]) # get price

    win.descendants(title='Организация:')[0].click_input()
    send_keys(comp1+'{ENTER}') # enter name1 into document
    time.sleep(timetosleep)
    win.descendants(title='Контрагент:')[0].click_input()
    send_keys(comp2+'{ENTER}{ENTER}') # enter name2 into document
    send_keys('{ENTER}')
    send_keys('{ENTER}')
    time.sleep(timetosleep)
    #                               from this
    send_keys('{ENTER}')
    send_keys('{ENTER}')
    send_keys('{ENTER}')
    send_keys('{INS}')
    send_keys('{DOWN}')
    send_keys('{ENTER}')
    #                               to this
    #                               i`m selecting place for entering the weight
    send_keys(weight[0]+weight[1]) # enter the weight
    send_keys('{RIGHT}')
    sent = ''

    for i in range(2, len(weight)):
        sent += weight[i]
    send_keys(sent)
    send_keys('{ENTER}')
    send_keys('{ENTER}')
    send_keys('{ENTER}')

    left = ''
    right = ''
    f = 0
    for i in range(0, len(price)): # formatting the price in needed format
        if price[i] == '.':
            f = 1
        else:
            if f == 0:
                left += price[i]
            else:
                right += price[i]
    # entering the price
    send_keys(left)
    send_keys('{RIGHT}')
    send_keys(right)
    send_keys('{ENTER}')
    send_keys('20 {ENTER}')

    time.sleep(timetosleep)
    send_keys('{ESC}')
    send_keys('{ESC}')

    win.descendants(title='Дополнительно')[0].click_input() # select advansed settings
    win.descendants(title="Адрес доставки:")[1].click_input() # select address to delete it

    time.sleep(timetosleep)
    send_keys('{DEL}') # deleting the address
    send_keys('{ENTER}')

    win.descendants(title="Склад:")[0].click_input() # select storage to delete it
    time.sleep(timetosleep)
    send_keys('^a {DEL}') # deleting it
    send_keys('Основний_'+comp1+'{ENTER}') # entering storage for company with name1

    win.descendants(title='Печать')[0].click_input() # select print to save the document
    send_keys('{ENTER}')
    send_keys('{DOWN}')
    send_keys('{DOWN}')
    send_keys('{ENTER}')
    time.sleep(timetosleep)
    send_keys('{ENTER}')

    win.descendants(title='Сохранить...')[0].click_input() # saving it
    send_keys(save_path) # enter the saving path
    send_keys('^{ENTER}')

    time.sleep(timetosleep)
    send_keys('{VK_ESCAPE}')

    win.descendants(title='Создать на основании')[0].click_input() # select other section for other document
    send_keys('{DOWN}')
    send_keys('{DOWN}')
    send_keys('{DOWN}')
    send_keys('{DOWN}')
    send_keys('{ENTER}')

    win.descendants(title='Документ\nрасчетов:')[0].click_input() # select new document
    send_keys('^a {DEL}')
    time.sleep(timetosleep)
    send_keys('{ENTER}')
    send_keys('{ENTER}')
    send_keys('{ENTER}')

    win.descendants(title='Печать')[0].click_input() # select print to save the document
    send_keys('{DOWN}')
    send_keys('{DOWN}')
    send_keys('{DOWN}')
    send_keys('{ENTER}')
    send_keys('{ENTER}')

    win.descendants(title='Сохранить...')[0].click_input() # saving it
    send_keys(save_path) # enter the saving path
    send_keys('^{ENTER}')
