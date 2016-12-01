from requests import get
from xml.dom import minidom
from random import shuffle

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from bs4 import BeautifulSoup
from dateutil.parser import *

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

import re


class Constants:
    XML_FEED = "https://www.coursera.org/sitemap~www~courses.xml"
    courses_urls = []
    key_names = ['Name', 'Language', 'Rating', 'Starts on', 'Duration (in weeks)']
    class_names = ['.title', '.language-info', '.ratings-text']
    all_courses = []
    xls_headerfill = PatternFill(start_color='E1E4D2', end_color='E1E4D2', fill_type='solid')
    xls_titlefont = Font(size=11, bold=True, color='000000')


def get_courses_list(amount=20, url_tag='loc'):

    try:
        xml_doc = minidom.parseString(get(Constants.XML_FEED).content)
        location_list = xml_doc.getElementsByTagName(url_tag)
        shuffle(location_list)

        Constants.courses_urls.extend([tag.firstChild.data
                                       for index, tag in enumerate(location_list) if index < amount])

    except Exception as error:
        print('An error occurred while getting courses xml:', error)

    return Constants.courses_urls


def get_course_info(course_slug):

    sleep_for = 30

    driver = webdriver.PhantomJS()
    driver.set_window_size(1124, 850)
    driver.get(course_slug)

    just_wait = lambda: WebDriverWait(driver, sleep_for).\
        until(EC.visibility_of_element_located((By.CLASS_NAME, "comfy")))

    print("fetching course: {}, [{}/{}]".format
         (course_slug, Constants.courses_urls.index(course_slug)+1, len(Constants.courses_urls)))

    while 1:
        try:
            just_wait()
        except BaseException:
            sleep_for += 10
            just_wait()
        else:
            html = driver.execute_script("return document.getElementById('rendered-content').innerHTML")
            driver.quit()
            return html


def parse_course(html):

    course = []
    soup = BeautifulSoup(html, "html.parser")

    get_element = lambda element: soup.select_one(element).get_text()

    try:
        for class_ in Constants.class_names:
            course.append(get_element(class_))
    except:
        course.append('Not rated yet')

    try:
        js_chunk = soup.find('script', type="application/ld+json").contents[0].strip()
        dates = re.findall("\d{4}-\d{2}-\d{2}", js_chunk)
    except:
        course.extend(['Date unknown', 'Duration unknown'])
    else:
        course.extend([dates[0], int((parse(dates[1])-parse(dates[0])).days/7-1)])

    Constants.all_courses.append(course)


def output_courses_info_to_xlsx(filepath):
    wb = Workbook()
    ws = wb.active
    ws.append(Constants.key_names)

    for course in Constants.all_courses:
        ws.append(course)

    for row in ws["A1:E1"]:
        for cell in row:
            cell.fill = Constants.xls_headerfill
            cell.font = Constants.xls_titlefont

    wb.save(filepath)


def main():

    for course in get_courses_list():
        parse_course(get_course_info(course))

    output_courses_info_to_xlsx("courses.xlsx")

if __name__ == '__main__':
    main()
