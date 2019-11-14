# import web driver
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from parsel import Selector
import requests
import time
import sys
import xlwt
import xlrd

def googleSearch():
    all_linkedin_links = []
    all_access_linkedin_links = []

    search_query = 'site:linkedin.com/in/ AND "sanitation"'
    driver.get('https://www.google.com')

    # Type search
    searchBar = driver.find_element_by_class_name('gLFyf')
    searchBar.send_keys(search_query)

    # Press search button
    search_button = driver.find_elements_by_name('btnK')
    search_button[1].click()

    current_target = 2
    nb_page_to_explore = 30

    for _ in range(nb_page_to_explore):
        # Re init all arrays
        this_all_linkedin_links = []
        this_linkedin_links = []
        this_access_linkedin_links = []

        this_all_linkedin_links = driver.find_elements_by_css_selector('.r > a')
        # Create a list purely made of the Linkedin link
        this_linkedin_links = [url.get_attribute("href") for url in this_all_linkedin_links if url.get_attribute("href").find('translate') == -1]

        # Merge new list found with the old ones
        all_linkedin_links = all_linkedin_links + this_linkedin_links

        # Target bottom page links to other google pages
        all_next_pages = driver.find_elements_by_xpath('//tbody/tr/td')

        for page in all_next_pages:
            try:
                if page.text and int(page.text) == current_target:
                    current_target += 1
                    page.click()
                    break
            except ValueError:
                print("Not number")

    driver.quit()

    return all_linkedin_links


def create_write_excel_file(all_linkedin_links):
     # Create a workbook and add a worksheet.
     wb = xlwt.Workbook()
     ws = wb.add_sheet('Scaper')
     ws.write(0,0, "Normal link")
     ws.write(0,1, "Access link")

     for i in range(len(all_linkedin_links)):
         ws.write(i+1, 0, all_linkedin_links[i])
         ws.write(i+1, 1, "https://translate.google.com/translate?hl=en&sl=fr&u=" + all_linkedin_links[i])

     wb.save('sanitation_results.xls')

def read_excel_file():
    book = xlrd.open_workbook('results.xls')
    sheet = book.sheet_by_index(0)

    for i in range(1, 1):
        try:
            driver.get(str(sheet.cell_value(i,1)))
            time.sleep(5)
            iframe = driver.find_elements_by_tag_name('iframe')[0]
            driver.switch_to_frame(iframe)
            sel = Selector(text=driver.page_source)
            name = sel.xpath('/html/body/main/section[1]/section/section[1]/div/div[1]/div[1]/h1/span/span/text()').get()
            job = sel.xpath('/html/body/main/section[1]/section/section[1]/div/div[1]/div[1]/h2/span/span/text()').get()
            about = sel.xpath('/html/body/main/section[1]/section/section[2]/p/text()').get()
            print(name, job, about)
        except:
            print("Page can't load")

    driver.quit()

if __name__ == "__main__":
    # specifies the path to the chromedriver.exe
    driver = webdriver.Chrome('/usr/local/bin/chromedriver')
    #all_linkedin_links = googleSearch()
    #create_write_excel_file(all_linkedin_links)
    read_excel_file()
