import csv
import ssl
import time

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import smtplib
from email.message import EmailMessage


def WebDriver():
    options = Options()
    options.add_experimental_option("excludeSwitches", ["ignore-certificate-errors"])
    options.add_argument('--disable-gpu')
    # options.add_argument('--headless')
    options.add_argument("--window-size=1920,1080")
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--allow-running-insecure-content')
    user_agent = 'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36'
    options.add_argument(f'user-agent={user_agent}')
    options.add_argument("--incognito")
    options.add_argument('disable-infobars')
    options.add_argument('--disable-browser-side-navigation')
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option('prefs', {'intl.accept_languages': 'en,en_US'})
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
    return driver


def send_mail(message):
    sender = 'atone0488@gmail.com'
    password = 'ooecriysztfuvfau'
    recevier = 'leoking627@gmail.com'
    em = EmailMessage()
    em['From'] = sender
    em['To'] = recevier
    em['Subject'] = 'appointment confirmation'
    em.set_content(message, subtype='html')
    # em.set_content(message)
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(sender, password)
        smtp.sendmail(sender, recevier, em.as_string())


def get_appointment():
    for row in csv.DictReader(open('appointmentdetail.csv', encoding='utf-8')):
        for profile_row in csv.DictReader(open('profles.csv', encoding='utf-8')):
            first_name = profile_row['First Name']
            last_name = profile_row['Last Name']
            dob = profile_row['DOB']
            driving_licence = profile_row['Driver License']
            state_driving_licence = profile_row['State - Driver License']
            email = profile_row['Email']
            phone_number = profile_row['Phone Number']

            driver = WebDriver()
            driver.maximize_window()
            url = 'https://nqa4.nemoqappointment.com/Booking/Booking/Index/fl34hs86'
            driver.get(url)
            driver.find_element(By.NAME, 'StartNextButton').click()
            driver.find_element(By.ID, 'AcceptInformationStorage').click()
            captcha_iframe = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'iframe')))
            ActionChains(driver).move_to_element(captcha_iframe).click().perform()

            captcha_box = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'g-recaptcha-response')))
            driver.execute_script("arguments[0].click()", captcha_box)
            time.sleep(20)

            driver.find_element(By.NAME, 'Next').click()

            select = Select(driver.find_element(By.ID, 'RegionId'))
            select.select_by_visible_text(row['country'])
            select = Select(driver.find_element(By.ID, 'SectionId'))
            select.select_by_visible_text(row['office'])
            time.sleep(5)
            select = Select(driver.find_element(By.ID, 'ServiceTypeId'))
            select.select_by_visible_text(row['appointment type'])

            # tomorrow_datetime = datetime.datetime.now() + datetime.timedelta(days=1)
            # tomorrow_date = tomorrow_datetime.date().strftime("%m/%d/%Y")

            date_ele = driver.find_element(By.ID, 'datepicker')
            driver.execute_script("arguments[0].value = '{}';".format(row['date']), date_ele)

            driver.find_element(By.NAME, 'TimeSearchButton').click()
            driver.implicitly_wait(10)
            try:
                driver.find_element(By.XPATH,
                                    "//*[contains(@headers,'{}')]//*[contains(@style,'background-color: #00FF00')][1]".format(
                                        str(row['date']).lstrip('0'))).click()
                driver.find_element(By.ID, 'booking-next').click()

                first_name_element = driver.find_element(By.ID, 'Customers_0__BookingFieldValues_0__Value')
                first_name_element.send_keys(first_name)

                last_name_element = driver.find_element(By.ID, 'Customers_0__BookingFieldValues_1__Value')
                last_name_element.send_keys(last_name)

                dob_element = driver.find_element(By.ID, 'Customers_0__BookingFieldValues_2__Value')
                dob_element.send_keys(dob)

                licence_no_element = driver.find_element(By.ID, 'Customers_0__BookingFieldValues_3__Value')
                licence_no_element.send_keys(driving_licence)

                state_licence_no_element = driver.find_element(By.ID,
                                                               'Customers_0__BookingFieldValues_4__Value')
                state_licence_no_element.send_keys(state_driving_licence)

                driver.find_element(By.NAME, 'Next').click()

                email_element = driver.find_element(By.ID, 'EmailAddress')
                email_element.send_keys(email)

                phone_element = driver.find_element(By.ID, 'PhoneNumber')
                phone_element.send_keys(phone_number)

                driver.find_element(By.ID, 'SelectedContacts_0__IsSelected').click()

                driver.find_element(By.ID, 'SelectedContacts_2__IsSelected').click()

                driver.find_element(By.NAME, 'Next').click()

                booked_time = driver.find_element(By.XPATH,
                                                  "//label[@for='BookedDateTime']/following-sibling::div/span").text

                department = driver.find_element(By.XPATH,
                                                 "//label[@for='BookedServiceGroupName']/following-sibling::div/span").text

                office = driver.find_element(By.XPATH,
                                             "//label[@for='BookedSectionName']/following-sibling::div/span").text

                appointment_type = driver.find_element(By.XPATH,
                                                       "//label[@for='BookedServiceTypeName']/following-sibling::div/span").text
                confirm_appointment = driver.find_element(By.XPATH, "//div[@class='btn-toolbar'][1]/input")
                confirm_appointment.click()

                confirmation_number = driver.find_element(By.XPATH,
                                                          "//label[@for='BookingNumber']/following-sibling::div/b").text
                message = """<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
                                            <html xmlns="http://www.w3.org/1999/xhtml">
                                            <head>
                                                <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
                                                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                                                <title>Appointment</title>
                                            </head>
                                            <body style="margin:0; border: none; background:#f5f5f5">
                                            <table align="center" border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
                                                <tr>
                                                    <td align="center" valign="top">
                                                        <table class="contenttable" border="0" cellpadding="0" cellspacing="0" width="600" bgcolor="#ffffff"
                                                               style="border-width: 8px; border-style: solid; border-collapse: separate; border-color:#ececec; margin-top:40px; font-family:Arial, Helvetica, sans-serif">
                                                            <tr>
                                                                <td>
                                                                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                                                        <tbody>
                                                                        <tr>
                                                                            <td width="100%" height="40">&nbsp;</td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td valign="top" align="center"><h2
                                                                                    style="font-size: 30px ; font-family: 'Waterfall', cursive ; text-align: center">
                                                                                Appointment Details </h2>
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td width="100%" height="40">&nbsp;</td>
                                                                        </tr>
                                                                        <tr>
                                                                        </tbody>
                                                                    </table>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td>
                                                                    <table border="1" cellpadding="2" cellspacing="1" width="100%"
                                                                           style="border-collapse: separate ; border:solid black 1px;border-radius:6px">
                                                                        <tbody>
                                                                        <tr class="table_row">
                                                                            <th class="table_head">Confirmation Number</th>
                                                                            <td class="table_value">{confirmation_number}</td>
                                                                        </tr>
                                                                        <tr class="table_row">
                                                                            <th class="table_head">First Name</th>
                                                                            <td class="table_value">{first_name}</td>
                                                                        </tr>
                                                                        <tr class="table_row">
                                                                            <th class="table_head">Last Name</th>
                                                                            <td class="table_value">{last_name}</td>
                                                                        </tr>
                                                                        <tr class="table_row">
                                                                            <th class="table_head">DOB</th>
                                                                            <td class="table_value">{dob}</td>
                                                                        </tr>
                                                                        <tr class="table_row">
                                                                            <th class="table_head">Email</th>
                                                                            <td class="table_value">{email}</td>
                                                                        </tr>
                                                                        <tr class="table_row">
                                                                            <th class="table_head">Phone Number</th>
                                                                            <td class="table_value">{phone_number}</td>
                                                                        </tr>
                                                                        <tr class="table_row">
                                                                            <th class="table_head">Booked Time</th>
                                                                            <td class="table_value">{booked_time}</td>
                                                                        </tr>
                                                                        <tr class="table_row">
                                                                            <th class="table_head">Department</th>
                                                                            <td class="table_value">{department}</td>
                                                                        </tr>
                                                                        <tr class="table_row">
                                                                            <th class="table_head">Office</th>
                                                                            <td class="table_value">{office}</td>
                                                                        </tr>
                                                                        <tr class="table_row">
                                                                            <th class="table_head">Appointment Type</th>
                                                                            <td class="table_value">{appointment_type}</td>
                                                                        </tr>
                                                                        <tr class="table_row">
                                                                            <th class="table_head">Driving Licence</th>
                                                                            <td class="table_value">{driving_licence}</td>
                                                                        </tr>
                                                                        <tr class="table_row">
                                                                            <th class="table_head">State Driving Licence</th>
                                                                            <td class="table_value">{state_driving_licence}</td>
                                                                        </tr>
                                                                        </tbody>
                                                                    </table>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#fcfcfc" class="tablepadding"
                                                                    style="padding:20px 0; border-top-width:1px;border-top-style:solid;border-top-color:#ececec;border-collapse:collapse">
                                                                    <table align="center">
                                                                    </table>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                </tr>
    
                                            </table>
                                            </body>
                                            </html>""".format(
                    confirmation_number=confirmation_number, first_name=first_name, last_name=last_name, dob=dob,
                    email=email,
                    phone_number=phone_number,
                    booked_time=booked_time, department=department, office=office,
                    appointment_type=appointment_type,
                    driving_licence=driving_licence, state_driving_licence=state_driving_licence)
                try:
                    send_mail(message)
                except Exception as e:
                    print(e)
                driver.quit()
            except:
                print(driver.find_element(By.XPATH, "//label[@style='color:Red;']/p").text)
                driver.quit()


if __name__ == '__main__':
    get_appointment()
    # schedule.every(60).seconds.do(get_appointment)
    # while True:
    #     schedule.run_pending()
    #     time.sleep(1)
