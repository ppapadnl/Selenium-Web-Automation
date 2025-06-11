from selenium import webdriver
from time import sleep
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.edge.options import Options
from datetime import timedelta
import base64
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
import glob
import os
import oracledb
import pandas as pd
import time
from pandasql import sqldf
from datetime import datetime

strDate = datetime.today().strftime('%d-%m-%Y')
current_month = datetime.now().month
current_year = datetime.now().year
desired_date = datetime.now() - timedelta(days=1)
desired_date_str = desired_date.strftime("%A, %B %d, %Y")
desired_date_str = desired_date_str.replace(" 0", " ")
current_date_chk = datetime.today()

if current_date_chk.weekday() == 6 or current_date_chk.weekday() == 0:
    print("Day is Sunday or Monday")
    quit()

if __name__ == "__main__":
    options = Options()
    options.add_argument("--force-device-scale-factor=0.8")
    browser = webdriver.Edge(options=options)
    browser.maximize_window()

    browser.get('https://se01.cos.prod.centiro.cloud/p16/insight/shipmentdetails/')

    try:
        myElem_1 = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'username')))
        myElem_1.send_keys("RPA_PYTHON@PP.COM")
        myElem_1.click()

        myElem_3 = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'password')))
        myElem_3.send_keys("Password")
        myElem_3.click()
        sleep(3)

        myElem_2 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "btnLogin")))
        myElem_2.click()
        sleep(2)

        myElem_4 = WebDriverWait(browser, 10).until(
            EC.element_to_be_clickable((By.XPATH,
                                        "/html/body/app-root/coordinator/ngds-layout/div/div[2]/action-bar-root/action-bar/table/thead/tr/th[2]/div/button")))
        myElem_4.click()
        sleep(1)

        myElem_4b = WebDriverWait(browser, 10).until(
            EC.element_to_be_clickable((By.XPATH,
                                        "/html/body/app-root/coordinator/ngds-layout/div/aside/div/filter-root/filter/form/div/div[4]")))
        myElem_4b.click()
        sleep(1)

        # Page Down on filter
        actions = ActionChains(browser)
        actions.send_keys(Keys.PAGE_DOWN).perform()
        # actions.send_keys(Keys.PAGE_DOWN).perform()

        myElem_4c = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/app-root/coordinator/ngds-layout/div/aside/div/filter-root/filter/form/div/ngds-filter-header/div/div[1]/button[2]")))
        myElem_4c.click()
        sleep(1)

        myElem_4d = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/app-root/coordinator/ngds-layout/div/aside/div/filter-root/filter/form/div/ngds-filter-header/div/div[2]/div/div/div/div[9]/div/div/label")))
        myElem_4d.click()
        myElem_4c.click()
        sleep(1)

        myElem_5 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='sender']/ngds-autocomplete-toggle/div/div[1]")))
        myElem_5.click()
        sleep(1)

        myElem_6 = WebDriverWait(browser, 10).until(EC.presence_of_element_located(
            (By.XPATH, "//*[@id='sender']/angular2-multiselect/div/div[2]/div[3]/div[2]/input")))
        myElem_6.send_keys("RWS TLB c/o WL Gore & Associates GMB")
        myElem_6.click()
        sleep(2)

        myElem_9 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='sender']/angular2-multiselect/div/div[2]/div[3]/div[4]/ul/div[2]/li/label")))
        myElem_9.click()
        sleep(1)
        myElem_4b.click()
        sleep(1)

        myElem_6New = WebDriverWait(browser, 10).until(EC.presence_of_element_located(
            (By.XPATH, "/html/body/app-root/coordinator/ngds-layout/div")))
        myElem_6New.click()
        sleep(2)
        # /html/body/app-root/coordinator/ngds-layout/div

        myElem_19 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
            (By.XPATH,
             "/html/body/app-root/coordinator/ngds-layout/div/aside/div/filter-root/filter/form/div/div[5]/div/form/div/div")))
        myElem_19.click()
        sleep(2)

        select_element1 = WebDriverWait(browser, 5).until(EC.element_to_be_clickable(
            (By.XPATH, "//select[@aria-label='Select month']"
             )))
        select = Select(select_element1)
        select.select_by_value(str(current_month))
        sleep(2)

        select_element3 = WebDriverWait(browser, 5).until(EC.element_to_be_clickable(
            (By.XPATH, "//select[@aria-label='Select year']"
             )))
        select = Select(select_element3)
        select.select_by_value(str(current_year))
        sleep(2)

        xpath_expression = f"//div[@aria-label='{desired_date_str}']"
        select_element2 = WebDriverWait(browser, 60).until(EC.element_to_be_clickable(
            (By.XPATH, xpath_expression)))
        select_element2.click()
        sleep(2)
        myElem_19.click()

        myElem_N20 = WebDriverWait(browser, 5).until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='shipmentStatus']/ngds-autocomplete-toggle/div/div[1]/div[1]")))
        myElem_N20.click()
        sleep(2)

        actions = ActionChains(browser)
        actions.send_keys(Keys.PAGE_DOWN).perform()

        myElem_N21 = WebDriverWait(browser, 15).until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='shipmentStatus']/angular2-multiselect/div/div[2]/div[3]/div[2]/ul/li[2]/input")))
        myElem_N21.click()
        sleep(2)

        myElem_N20 = WebDriverWait(browser, 5).until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='shipmentStatus']/ngds-autocomplete-toggle/div/div[1]/div[1]")))
        myElem_N20.click()
        sleep(3)

        myElem_N22 = WebDriverWait(browser, 5).until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='shipmentStatus']/angular2-multiselect/div/div[2]/div[3]/div[2]/ul/li[3]/input")))
        myElem_N22.click()
        sleep(2)

        myElem_20 = WebDriverWait(browser, 5).until(EC.element_to_be_clickable(
            (By.XPATH,
             "/html/body/app-root/coordinator/ngds-layout/div/aside/div/filter-root/filter/form/div/ngds-filter-header/div/button"
             )))
        myElem_20.click()
        sleep(2)

        myElem_21 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
            (By.XPATH,
             "/html/body/app-root/coordinator/ngds-layout/div/div[2]/div/shipment-table-root/shipment-table/div/table/thead/tr/th[1]/div/input"
             )))
        myElem_21.click()
        sleep(1)

        myElem_23 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
            (By.XPATH,
             "/html/body/app-root/coordinator/ngds-layout/div/div[2]/action-bar-root/action-bar/table/thead/tr/th[1]/div[5]/ngds-dropdown-menu/div/div/button")))
        myElem_23.click()
        sleep(1)

        myElem_24 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
            (By.XPATH,
             "/html/body/app-root/coordinator/ngds-layout/div/div[2]/action-bar-root/action-bar/table/thead/tr/th[1]/div[5]/ngds-dropdown-menu/div/div[2]/ngds-dropdown-menu-item[1]/button")))
        myElem_24.click()
        sleep(2)

        myElem_25 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='fileExport']")))
        myElem_25.click()
        sleep(7)

        # Connection to instant-client to connect to the database:
        oracledb.init_oracle_client(
            lib_dir=r"C:\instantclient_21_9")
        oracledb.clientversion()

        # WMS production 2019:
        username = "USER_NAME"
        password = "PASSWORD"
        service_name = "SRVCNAME"
        port = 1521
        host = "HOST"

        # Create a connection the WMS production database:
        conn_wmsdb = oracledb.connect(user=username, password=password, port=port, service_name=service_name, host=host)

        # SQL query for order_header last 7 days for WLG to find the order_reference and purchase_order
        query_oh = """select
        creation_date,
        client_id,
        order_id,
        purchase_order,
        order_reference
        FROM dcsdba.order_header
        where client_id = 'WLG'
        and creation_date >= SYSDATE - INTERVAL '7' DAY 
        order by order_id
        """

        # Create dataframe for order_reference and purchase_order
        df_oracle = pd.read_sql(query_oh, con=conn_wmsdb).astype(str)

        # Check if empty or not
        if df_oracle.empty:
            print("No data found on SQL Query, close script.")
            exit()
        else:
            print("df_oracle created")

        time.sleep(1)
        conn_wmsdb.close()

        # Create dataframe from the TMS downloaded file:
        centiro_folder_path = r"C:\Users\Pypower\Downloads"
        files = [os.path.join(centiro_folder_path, file) for file in os.listdir(centiro_folder_path) if
                 file.lower().endswith('.xlsx')]
        files.sort(key=os.path.getmtime, reverse=True)
        if files:
            centiro_file_input = files[0]
            df_centiro = pd.read_excel(centiro_file_input)
            print(f"Excel file: {centiro_file_input}")

        # Consolidate df_centiro with df_oracle to create wlg_report:
        report = (sqldf('''SELECT 
        CEN.[Order No], 
        WMS.[order_reference] as [Order Reference], 
        WMS.[purchase_order] as [Purchase Order],
        CEN.*
        FROM df_centiro CEN
        LEFT JOIN df_oracle WMS on WMS.order_id = CEN.[Order NO]
        '''))
        report = report.loc[:, ~report.columns.duplicated()]

        # Save wlg_report to excel with current date in filename:
        directory = r"C:\Users\NLTLG03.pypower\Downloads"
        filename = f"ExportedShipments_{strDate}.xlsx"
        file_path = os.path.join(directory, filename)
        wlg_report.to_excel(file_path, index=False)
        print("New report is saved as:", filename)

        download_folder = r"C:\Users\NLTLG03.pypower\Downloads"
        files = glob.glob(os.path.join(download_folder, '*'))
        files.sort(key=os.path.getmtime, reverse=True)
        most_recent_file = files[0]
        filename = most_recent_file
        image_file = open(r"C:\Users\Pypower\Documents\PythonProjects\pyReportAuto_Oracle\pic.png",
                          'rb').read()
        encoded_image = base64.b64encode(image_file).decode("utf-8")
        from_email = "Donotreply@nl.smtp.com"
        to_emails = ["recipient.tilburg@n.com", "recipient@gmail.com", "recipient@gmail.com",
                     "recipient@gmail.com"]
        cc_emails = ["recipient@gmail.com", recipient@gmail.com", "recipient@gmail.com"]
        subject = "Automated Notification - Report SENDER REF from yesterday Shipments    " + desired_date_str
        message = '<p3>Dear Customer Service,  </p3>' '<br>' '<br>' '<br>' \
                  '<p3>Please note that this is an automated email, and there is no' \
                  ' need to reply. </p3>' '<br>''<br>''<p3>We wanted to inform you that a report has been triggered for ' \
                  'CLIENT, covering multiple Carries and service levels from Shipments that were shipped on {}.</p3>''<br>' \
                  '<br>' \
                  '<p3>The attached file is the report generated by our Transport Management System (TMS).</p3>''<br>' '<br>' \
                  '<p3>If you have any questions or concerns regarding this report or require further assistance, we kindly ' \
                  'request that you reach out to our customer service or your designated Account Manager.</p3>''<br>''<br>''<br>' \
                  '<p3>Best regards,</p3>' \
                  '<h5 style="color: red">Rhenus IT</h5>' \
                  '<img src="data:image/png;base64,%s"/>'.format(desired_date_str) % encoded_image

        attachmentEx = open(filename, 'rb')
        attachment_filename = os.path.basename(filename)
        msg = MIMEMultipart()
        msg['To'] = ", ".join(to_emails)
        msg['Cc'] = ", ".join(cc_emails)
        msg['Subject'] = subject
        msg.attach(MIMEText(message, 'html'))
        attachment = MIMEBase('application', 'octet-stream')
        attachment.set_payload(attachmentEx.read())
        encoders.encode_base64(attachment)
        attachment.add_header('Content-Disposition', f'attachment; filename={attachment_filename}')
        msg.attach(attachment)

        server = smtplib.SMTP('smtp.service.xxx.dd', 22) # SMTP & Port

    except TimeoutException:
        print("No element found")
        browser.quit()
        server.quit()
        exit()
    sleep(1)
    server.sendmail(from_email, to_emails + cc_emails, msg.as_string())
    print("mail sent")
    sleep(2)
    server.quit()
    browser.quit()
