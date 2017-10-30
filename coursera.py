import requests
from bs4 import BeautifulSoup
import random
from openpyxl import Workbook
import argparse
import re


XML_COURSE_LIST = 'https://www.coursera.org/sitemap~www~courses.xml'
COURSE_COUNT = 10


def get_courses_list():
    response = requests.get(XML_COURSE_LIST)
    soup = BeautifulSoup(response.text, 'lxml')
    course_list = [course.text for course in soup.body.find_all('loc')]
    random_course = random.sample(course_list, COURSE_COUNT)
    return random_course


def get_course_info(course_url):
    response = requests.get(url=course_url)
    soup = BeautifulSoup(response.text.encode('utf-8'), 'html.parser')
    course_name = soup.find('h1', class_='title display-3-text')
    if course_name:
        course_name = course_name.text
    else:
        course_name = None
    course_lang = soup.find('div', class_='rc-Language')
    if course_lang:
        course_lang = course_lang.text
    else:
        course_lang = None
    assesment = soup.find('div',
                          class_='rc-RatingsHeader horizontal-box'
                                 ' align-items-absolute-center')
    if assesment:
        assesment = re.search(r'[\d.]+', assesment.text).group()
    else:
        assesment = None
    start_date = soup.find('div',
                           class_='startdate rc-StartDateString'
                                  ' caption-text').text
    course_duration = soup.find_all('div',
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


def output_courses_info_to_xlsx(filepath, courses_info):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "course_info"
    ws1.append(['Course name',
                'Course lang',
                'Assesment',
                'Start date',
                'Course duration in weeks'])
    for course_info in courses_info:
        exel_row = [course_info['course_name'],
                    course_info['course_lang'],
                    course_info['assesment'],
                    course_info['start_date'],
                    course_info['course_duration']]
        ws1.append(exel_row)
    wb.save(filename=filepath)


def create_parser():
    parser = argparse.ArgumentParser(description='course_info')
    parser.add_argument("output_file", nargs='?', const=1,
                        default='book.xlsx',
                        type=str, help="path to output file")
    return parser


if __name__ == '__main__':
    courses_info = []
    parser = create_parser()
    args = parser.parse_args()
    for course_url in get_courses_list():
        courses_info.append(get_course_info(course_url))
    output_courses_info_to_xlsx(args.output_file, courses_info)
