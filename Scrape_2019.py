import pandas as pd

import xlsxwriter

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException

chromedriver = "C:\\Users\\Matija\\Downloads\\chromedriver_win32\\chromedriver"

browser = webdriver.Chrome(chromedriver)
browser.get('https://Username:Password@avr.fringeworld.com.au/media_information')

page_links = browser.find_elements_by_xpath("//span[@class='page']/a")

page_links_text = []
page_links_text.append('https://Username:Password@avr.fringeworld.com.au/media_information')

for x in page_links:
    page_links_text.append(x.get_attribute('href'))

show_names = []
categories = []
presenter_names = []
presenter_phones = []
presenter_emails = []
contact_names = []
contact_phones = []
contact_emails = []

for i in page_links_text:
    try:
        browser.get(i)
        values_element = browser.find_elements_by_xpath("//div[@class='col-md-4']/a")
        links = [x.get_attribute('href') for x in values_element]
        for j in links:
            browser.get(j)
            show_name = browser.find_element_by_xpath("//span[@class='event-name']")
            show_genre = browser.find_element_by_xpath("//span[@class='genres']")

            try:
                presenter_name = browser.find_element_by_xpath("//div[@class='information spacing--xxx-tight']/p/b").text
            except NoSuchElementException:
                presenter_name = "N/A"
            try:
                presenter_phone = browser.find_element_by_xpath("//div[@class='information spacing--xxx-tight']/p/a[1]").get_attribute('href')
            except NoSuchElementException:
                presenter_phone = "N/A"
            try:
                presenter_email = browser.find_element_by_xpath("//div[@class='information spacing--xxx-tight']/p/a[2]").get_attribute('href')
            except NoSuchElementException:
                presenter_email = "N/A"
            try:
                contact_name = browser.find_element_by_xpath("//ul/li/b").text
            except NoSuchElementException:
                contact_name = "N/A"
            try:
                contact_details = browser.find_element_by_xpath("//ul/li").text.split('\n')
                contact_name = contact_details[0]
                contact_phone = contact_details[1]
            except NoSuchElementException:
                    contact_name, contact_phone = "N/A", "N/A"
            except IndexError:
                contact_name, contact_phone = "N/A", "N/A"
            try:
                contact_email = browser.find_element_by_xpath("//ul/li/a[2]").get_attribute("href")
            except NoSuchElementException:
                contact_email = "N/A"

            show_names.append(show_name.text)
            categories.append(show_genre.text)
            presenter_names.append(presenter_name)
            presenter_phones.append(presenter_phone)
            presenter_emails.append(presenter_email)
            contact_names.append(contact_name)
            contact_phones.append(contact_phone)
            contact_emails.append(contact_email)
    except TimeoutException:
        print("Webpage timed out")

spreadsheet = {'Show Name': show_names,
               'Genre / Subgenre': categories,
               'Presenter Contact': presenter_names,
               'Presenter Phone': presenter_phones,
               'Presenter Email': presenter_emails,
               'Publicity Contact': contact_names,
               'Publicity Phone': contact_phones,
               'Publicity Email': contact_emails}

df = pd.DataFrame.from_dict(spreadsheet)

df = df[['Show Name', 'Genre / Subgenre', 'Presenter Contact', 'Presenter Phone', 'Presenter Email', 'Publicity Contact', 'Publicity Phone', 'Publicity Email']]

writer = pd.ExcelWriter('Fringe_Contacts.xlsx', engine='xlsxwriter')
df.to_excel(writer, index=False)
writer.close()