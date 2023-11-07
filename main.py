from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import xlsxwriter
from selenium import webdriver

driver = webdriver.Chrome(executable_path=ChromeDriverManager().install())

workbook = xlsxwriter.Workbook('c:\\testtt.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': True})
worksheet.write('A1', 'نام فارسی', bold)
worksheet.write('B1', 'کد ژنریک', bold)
worksheet.write('C1', 'نام', bold)
worksheet.write('D1', 'نام عمومی', bold)
worksheet.write('E1', 'شکل دارویی', bold)
worksheet.write('F1', 'نحوه مصرف', bold)
worksheet.write('G1', 'صاحب پروانه', bold)
worksheet.write('H1', 'برند', bold)
worksheet.write('I1', 'تولید کننده', bold)
worksheet.write('J1', 'قیمت مصرف کننده', bold)
worksheet.write('K1', 'قیمت واحد', bold)
worksheet.write('L1', 'GTIN', bold)
worksheet.write('M1', 'IRC', bold)
worksheet.write('N1', 'تعداد در بسته', bold)
worksheet.write('O1', 'ATC', bold)
worksheet.write('P1', 'تاریخ اعتبار پروانه', bold)
worksheet.write('Q1', 'پروانه فوریتی', bold)

tbl = []
row = 1
col = 0
# driver.maximize_window()
# driver=webdriver.Chrome() 
# let_entries = ['cap', 'Topical', 'injection', 'drops', 'ointment', 'powder', 'Vaccine', 'Suspension', 'Syrup', 'Elixir', 'Shampoo'
#     , 'Rectal', 'Gel', 'Lotion', 'Suppository', 'Cream', 'liquid', 'INHAL',
#     'Enema', 'ORALLY', 'SPRAY', 'ARTIFICIAL', 'TEARS', 'vaginal', 'RESPIRATORY', 'Irrigation', 'tab']
let_entries = ['tab']

m = 0
n = 1
r = 1
MM = 3
MN = 3

while True:
    try:
        for i in range(m, len(let_entries), 1):
            term = let_entries[i]
            wait = WebDriverWait(driver, 60)
            if (n == 1) & (r == 1):
                page_url = "https://irc.fda.gov.ir/NFI/Search?Term=" + term + "&PageNumber=1&PageSize=1"
                driver.get(page_url)
                wait.until(EC.element_to_be_clickable(
                    (By.XPATH, '/html/body/div[1]/div[1]/nav/ul/li[7]/a'))).send_keys(Keys.ENTER)
                All_Element = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, '/html/body/div[1]/div[1]/nav/ul/li[6]/a')))
                All_ElementOfPerKeword = int(All_Element.text)
                # Quantity page of List element
                NN = (All_ElementOfPerKeword // 2)
                mode = All_ElementOfPerKeword % 2
                if mode == 0:
                    NN = NN + 1
                else:
                    NN = NN + 2
                tbl.append(All_ElementOfPerKeword)
                print('NN:', NN)
                print('All_ElementOfPerKeword=', All_ElementOfPerKeword)

            for ii in range(n, NN, 1):
                print('ii=', ii)
                page_url = "https://irc.fda.gov.ir/NFI/Search?Term=" + \
                           str(term) + "&PageNumber=" + str(ii) + "&PageSize=2"
                # fetching the Number of Opened tabs
                driver_len = len(driver.window_handles)
                #                 print("Length of Driver,before1 = ", driver_len)
                driver.get(page_url)
                p = driver.current_window_handle
                # obtain parent window handle
                parent = driver.window_handles[0]
                #                 print('parent')

                # fetching the Number of Opened tabs
                driver_len = len(driver.window_handles)
                #                 print("Length of Driver,before2 = ", driver_len)
                if (ii == NN - 1):
                    n = 1
                    if mode == 0:
                        MM = 3
                        MN = 3
                    else:
                        MM = mode + 1
                        MN = mode + 1

                for iii in range(r, MM, 1):
                    # fetching the Number of Opened tabs
                    driver_len = len(driver.window_handles)
                    #                     print("Length of Driver,before3 = ", driver_len)
                    try:
                        title_Fa1 = wait.until(EC.element_to_be_clickable(
                            (By.XPATH,
                             '/html/body/div[1]/form[2]/div[2]/div/div/div[2]/table/tbody/tr[{}]/td/div[1]/div[2]/div[1]/span[1]/a'.format(
                                 iii))))
                        title_Fa = title_Fa1.text
                    except:
                        title_Fa = 'Not found!'

                    try:
                        Generic = wait.until(EC.element_to_be_clickable(
                            (By.XPATH,
                             '/html/body/div[1]/form[2]/div[2]/div/div/div[2]/table/tbody/tr[{}]/td/div[1]/div[2]/div[3]/div[3]/span'.format(
                                 iii))))
                        Generic_1 = Generic.text
                    except:
                        try:
                            Generic = wait.until(EC.element_to_be_clickable(
                                (By.XPATH,
                                 '/html/body/div[1]/form[2]/div[2]/div/div/div[2]/table/tbody/tr[{}]/td/div[1]/div[2]/div[3]/div[2]/span'.format(
                                     iii))))
                            Generic_1 = Generic.text
                        except:
                            Generic_1 = 'Not found!'

                    segtit1 = wait.until(EC.element_to_be_clickable(
                        (By.XPATH,
                         '/html/body/div[1]/form[2]/div[2]/div/div/div[2]/table/tbody/tr[{}]/td/div[1]/div[2]/div[1]/span[1]/a'.format(
                             iii))))
                    action = ActionChains(driver)
                    action.key_down(Keys.CONTROL).click(
                        segtit1).key_up(Keys.CONTROL).perform()

                    # obtain browser tab window
                    chld = driver.window_handles[1]
                    driver.switch_to.window(chld)

                    # fetching the Number of Opened tabs
                    driver_len = len(driver.window_handles)
                    #                     print("Length of Driver,next = ", driver_len)

                    # مشخصات
                    name = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div[1]/div/div/div[3]/div[1]/div[1]/span')))
                    Public_name = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div[1]/div/div/div[3]/div[1]/div[2]/bdo')))
                    shekldaroii = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div[1]/div/div/div[3]/div[2]/div[1]/span')))
                    Nahve_Masraf = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div[1]/div/div/div[3]/div[2]/div[2]/bdo')))
                    parvane = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div[1]/div/div/div[3]/div[3]/div[1]/span')))
                    brand = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div[1]/div/div/div[3]/div[3]/div[2]/span')))
                    tolid = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div[1]/div/div/div[3]/div[4]/div[1]/span')))
                    price_Masraf = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div[1]/div/div/div[3]/div[5]/div[1]/span[1]')))
                    price_Unit = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div[1]/div/div/div[3]/div[5]/div[2]/span[1]')))
                    GTIN = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div[1]/div/div/div[3]/div[6]/div[1]/span')))
                    IRC = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div[1]/div/div/div[3]/div[6]/div[2]/span')))
                    NoINPack = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div[1]/div/div/div[3]/div[7]/div[1]/bdo')))
                    ATC = wait.until(EC.element_to_be_clickable(
                        (By.XPATH,
                         '/html/body/div[1]/div[5]/div[1]/div[1]/div/div/div[3]/div[9]/div/div/div[2]/div[3]/label')))
                    Credit_Date = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div[1]/div/div/div[3]/div[4]/div[2]/span')))
                    Parvane_Foriati = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div[1]/div/div/div[3]/div[7]/div[2]/span')))

                    worksheet.write(row, col, title_Fa)
                    worksheet.write(row, col + 1, Generic_1)
                    worksheet.write(row, col + 2, name.text.strip())
                    worksheet.write(row, col + 3, Public_name.text.strip())
                    worksheet.write(row, col + 4, shekldaroii.text.strip())
                    worksheet.write(row, col + 5, Nahve_Masraf.text.strip())
                    worksheet.write(row, col + 6, parvane.text.strip())
                    worksheet.write(row, col + 7, brand.text.strip())
                    worksheet.write(row, col + 8, tolid.text.strip())
                    worksheet.write(row, col + 9, price_Masraf.text.strip())
                    worksheet.write(row, col + 10, price_Unit.text.strip())
                    worksheet.write(row, col + 11, GTIN.text.strip())
                    worksheet.write(row, col + 12, IRC.text.strip())
                    worksheet.write(row, col + 13, NoINPack.text.strip())
                    worksheet.write(row, col + 14, ATC.text.strip())
                    worksheet.write(row, col + 15, Credit_Date.text.strip())
                    worksheet.write(row, col + 16, Parvane_Foriati.text.strip())
                    driver.close()
                    driver.switch_to.window(parent)

                    row += 1

                    if iii == MM - 1:
                        r = 1

                    if (ii == NN - 1) & (iii == MM - 1):
                        MM = 3

        if (i == len(let_entries) - 1) & (ii == NN - 1) & (iii == MN - 1):
            workbook.close()
            print(tbl)
            break

    except:
        m = i
        n = ii
        r = iii
        NN = NN
        mode = mode
        if driver_len >= 2:
            driver.close()
            driver.switch_to.window(parent)
