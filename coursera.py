import requests
from bs4 import BeautifulSoup
import random
from openpyxl import Workbook
import argparse
import re


XML_COURSE_LIST = 'https://www.coursera.org/sitemap~www~courses.xml'
COURSE_COUNT = 20


def get_courses_list():
    response = requests.get(XML_COURSE_LIST)
    soup = BeautifulSoup(response.text, 'lxml')
    course_list = [course.text for course in soup.body.find_all('loc')]
    return course_list


def get_random_course(course_list):
    random_course = random.sample(course_list, COURSE_COUNT)
    return random_course


def get_course_page(course_url):
    response = requests.get(url=course_url)
    response.encoding = 'UTF-8'
    html = BeautifulSoup(response.text, 'html.parser')
    return html


def get_course_info(html):
    course_name = html.find('h1', class_='title display-3-text')
    if course_name:
        course_name = course_name.text
    else:
        course_name = None
    course_lang = html.find('div', class_='rc-Language')
    if course_lang:
        course_lang = course_lang.text
    else:
        course_lang = None
    assesment = html.find('div',
                          class_='rc-RatingsHeader horizontal-box'
                                 ' align-items-absolute-center')
    if assesment:
        assesment = re.search(r'[\d.]+', assesment.text).group()
    else:
        assesment = None
    start_date = html.find('div',
                           class_='startdate rc-StartDateString'
                                  ' caption-text')
    if start_date:
        start_date = re.search(r'\w+ \d+', start_date.text).group()
    else:
        start_date = None
    course_duration = html.find_all('div',
                                    class_='week-heading body-2-text')
    if course_duration:
        course_duration = course_duration[-1].text[5:]
    else:
        course_duration = None
    return {'course_name': course_name,
            'course_lang': course_lang,
            'assesment': assesment,
            'start_date': start_date,
            'course_duration': course_duration}


def create_table_content(courses_info):
    table_head = [['Course name',
                   'Course lang',
                   'Assesment',
                   'Start date',
                   'Course duration in weeks']]
    table_content = [[course_info['course_name'],
                      course_info['course_lang'],
                      course_info['assesment'],
                      course_info['start_date'],
                      course_info['course_duration']]
                     for course_info in courses_info]
    table = table_head + table_content
    return table


def output_courses_info_to_xlsx(filepath, table):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "course_info"
    ws1.column_dimensions['A'].width = 30.0
    ws1.column_dimensions['B'].width = 10.0
    ws1.column_dimensions['C'].width = 10.0
    ws1.column_dimensions['D'].width = 10.0
    ws1.column_dimensions['E'].width = 20.0
    table_head, table_content = table[0], table[1::]
    ws1.append(table_head)
    for exel_row in table_content:
        ws1.append(exel_row)
    wb.save(filename=filepath)


def create_parser():
    parser = argparse.ArgumentParser(description='course_info')
    parser.add_argument("output_file", nargs='?', const=1,
                        default='book.xlsx',
                        type=str, help="path to output file")
    return parser


if __name__ == '__main__':
    parser = create_parser()
    args = parser.parse_args()
    courses_info = [get_course_info(get_course_page(course_url))
                    for course_url in get_random_course(get_courses_list())]
    output_courses_info_to_xlsx(args.output_file,
                                create_table_content(courses_info))
