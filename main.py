from selenium import webdriver
from time import sleep
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.edge.options import Options
from paaswordz import passo
from PIL import Image
from datetime import datetime
import base64
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
from datetime import timedelta

strDate = datetime.today().strftime('%d-%m-%Y')


def work_days(start_date, working_days):
    current_date = start_date
    while working_days > 0:
        current_date -= timedelta(days=1)
        if current_date.weekday() < 5:
            working_days -= 1
    return current_date


starting_date = datetime.now()
desired_date = work_days(starting_date, 5)
current_month = desired_date.month
current_year = desired_date.year
desired_date_str = desired_date.strftime("%A, %B %d, %Y")
desired_date_str = desired_date_str.replace(" 0", " ")

if __name__ == "__main__":
    options = Options()
    options.add_argument("--force-device-scale-factor=0.8")
    browser = webdriver.Edge(options=options)
    browser.maximize_window()

    browser.get('https://TMS.cloud/u01/tms/mastershipments/close')

    try:
        myElem_1 = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'username')))
        myElem_1.send_keys("TST_RPA_PYTHON@FMAIL.COM")
        myElem_1.click()

        myElem_3 = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'password')))
        myElem_3.send_keys(passo)
        myElem_3.click()
        sleep(2)

        myElem_2 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, "btnLogin")))
        myElem_2.click()
        sleep(2)

        myElem_4 = WebDriverWait(browser, 10).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/div/close"
                                                  "-shipment/close-shipment-view/ngds-layout/div/div[1]/div/button")))
        myElem_4.click()
        sleep(1)
        # Scroll
        browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        print("Here o.K")

        myElem_5 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/app-root/div/close-shipment/close-shipment-vie"
                       "w/ngds-layout/div/aside/filter-sidebar/form/div/div[4]/div/ngds-autocomplete-search-single/ngds-autocomplete-toggle/div/div[1]")))
        myElem_5.click()
        sleep(1)

        myElem_6 = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/app-ro"
                                                                                              "ot/div/close-shipment/close-shipment-view/ngds-layout/div/aside/filter-sidebar/form/div/div[4]/div/ngds-autocomplete-search-single/angular2-multiselect/div/div[2]/div[3]/div[1]/input")))
        myElem_6.send_keys("DHL Parcel")
        myElem_6.click()
        sleep(1)

        myElem_9 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
            (By.XPATH,
             "/html/body/app-root/div/close-shipment/close-shipment-view/ngds-layout/di"
             "v/aside/filter-sidebar/form/div/div[4]/div/ngds-autocomplete-search-single/angular2-multiselect/div/div[2]/div[3]/div[2]/ul/div[2]/li[2]")))
        myElem_9.click()
        sleep(1)

        myElem_new5 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
            (By.XPATH,
             "/html/body/app-root/div/close-shipment/close-shipment-view/ngds-layout/div/aside/filter-sidebar/form/div/div[3]/div/ngds-autocomplete-search-multi/ngds-autocomplete-toggle/div/div[1]/div[1]")))
        myElem_new5.click()
        sleep(1)

        myElem_new6 = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,
                                                                                       "/html/body/app-root/div/close-shipment/close-shipment-view/ngds-layout/div/aside/filter-sidebar/form/div/div[3]/div/ngds-autocomplete-search-multi/angular2-multiselect/div/div[2]/div[3]/div[2]/input")))
        myElem_new6.send_keys(
            "RWS TLB c/o Raymarine UK")  # "RWS TLB c/o Welch Allyn Ltd."  "RWS TLB c/o Penumbra" "RWS TLB c/o KCI Manufacturing"
        myElem_new6.click()
        sleep(2)
        print("kalo edo1")

        myElem_new9 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
            (By.XPATH,
             "/html/body/app-root/div/close-shipment/close-shipment-view/ngds-layout/div/aside/filter-sidebar/form/div/div[3]/div/ngds-autocomplete-search-multi/angular2-multiselect/div/div[2]/div[3]/div[3]/ul/div[2]/li/label")))
        myElem_new9.click()
        myElem_new5.click()
        sleep(1)
        print("kalo edo2 perimene")

        myElem_DateClick = WebDriverWait(browser, 7).until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='created-date-from']")))
        myElem_DateClick.click()
        sleep(1)

        select_element1 = WebDriverWait(browser, 5).until(EC.element_to_be_clickable(
            (By.XPATH, "//select[@aria-label='Select month']"
             )))
        select = Select(select_element1)
        select.select_by_value(str(current_month))
        sleep(1)

        select_element3 = WebDriverWait(browser, 5).until(EC.element_to_be_clickable(
            (By.XPATH, "//select[@aria-label='Select year']"
             )))
        select = Select(select_element3)
        select.select_by_value(str(current_year))
        sleep(1)

        xpath_expression = f"//div[@aria-label='{desired_date_str}']"
        select_element2 = WebDriverWait(browser, 60).until(EC.element_to_be_clickable(
            (By.XPATH, xpath_expression)))
        select_element2.click()
        sleep(1)

        myElem_9 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
            (By.XPATH,
             "/html/body/app-root/div/close-shipment/close-shipment-view/ngds-layout/div/aside/filter-sidebar/form/ngds-filter-header/div/button")))
        myElem_9.click()
        sleep(1)

        myElem_Checkboxez = browser.find_elements(By.XPATH, "//td//input[@type='checkbox']")
        for checkbox in myElem_Checkboxez:
            checkbox.click()
            sleep(1)

        # sleep(1)
        browser.save_screenshot("Masters.png")
        image = Image.open(r"C:\Users\SRV\PycharmProjects\pythonProject\Masters.png")  # get image
        img_convert = image.convert('RGB')
        img_convert.save(r"C:\Users\SRV\PycharmProjects\pythonProject\Masters.pdf")  # save pdf

        image_file = open("C:/Users/SRV/PycharmProjects/MAILSMTP1/Logo.png", 'rb').read()
        encoded_image = base64.b64encode(image_file).decode("utf-8")

        from_email = "TSTDonotreply@nl.rhenus.com"
        to_emails = ["receipient1@fmail.com", "receipient2@fmail.com", "receipient3@fmail.com",
                     "receipient4@fmail.com"]
        cc_emails = ["receipient_cc@fmail.com"]

        subject = "Test Automated Notification - Manifest for DHL_BBE    " + strDate
        message = '<p3>Dear Customer Service,  </p3>' '<br>' '<br>' '<br>' \
                  '<p3>Please note that this is a <b>Test</b> automated email, and there is no' \
                  ' need to reply. </p3>' '<br>''<br>''<p3>We wanted to inform you that a manifest has been performed for the ' \
                  'carrier ID: DHL Parcel, covering BBE service levels for Client XXX.</p3>''<br>' \
                  '<br>' \
                  '<p3>Attached to this email is a screenshot displaying the shipments for which the End of Day (EOD) and ' \
                  'manifest processes were successfully completed.</p3>' \
                  '<p3>If you have any questions or concerns regarding this manifest or require further assistance, we kindly ' \
                  'request that you raise a ticket with the Jira helpdesk.</p3>''<br>''<br>''<br>' \
                  '<p3>Best regards,</p3>' \
                  '<h5 style="color: red"> IT support</h5>' \
                  '<img src="data:image/png;base64,%s"/>' % encoded_image

        msg = MIMEMultipart()
        msg['To'] = ", ".join(to_emails)
        msg['Cc'] = ", ".join(cc_emails)
        msg['Subject'] = subject
        msg.attach(MIMEText(message, 'html'))
        pdf_filename = r"Masters.pdf"
        pdf_attachment = open(pdf_filename, "rb")
        attachment = MIMEBase('application', 'octet-stream')
        attachment.set_payload(pdf_attachment.read())
        encoders.encode_base64(attachment)
        attachment.add_header('Content-Disposition', f'attachment; filename={pdf_filename}')
        msg.attach(attachment)

        pdf_attachment.close()

        server = smtplib.SMTP('smtp.service.srv.ca', 25)

        myElem_20 = WebDriverWait(browser, 5).until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/app-root/div/close-shipment/close-shipment-view/ngds-layout/div/div[1]/button"
             )))
        myElem_20.click()
        sleep(3)

        myElem_21 = WebDriverWait(browser, 10).until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/ngb-modal-window/div/div/div/button[1]"
             )))
        myElem_21.click()
        sleep(1)

    except TimeoutException:
        print("No element found")
        browser.quit()
        exit()
    sleep(1)
    server.sendmail(from_email, to_emails + cc_emails, msg.as_string())
    print("mail sent")
    sleep(1)
    browser.quit()
    server.quit()

