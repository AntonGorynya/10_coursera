import requests
from bs4 import BeautifulSoup
import random
from openpyxl import Workbook


XML_COURSE_LIST = 'https://www.coursera.org/sitemap~www~courses.xml'
COURSE_COUNT = 20


def get_courses_list():
    response = requests.get(XML_COURSE_LIST)
    soup = BeautifulSoup(response.text, 'lxml')
    course_list = [course.text for course in soup.body.find_all('loc')]
    random_course = random.sample(course_list, COURSE_COUNT)
    return random_course


def get_course_info(course_url):
    response = requests.get(course_url)
    soup = BeautifulSoup(response.text, 'html.parser')
    course_name = soup.find('h1', class_='title display-3-text').text
    course_lang = soup.find('div', class_='rc-Language').text
    assesment = soup.find('div',
                          class_='rc-RatingsHeader horizontal-box'
                                 ' align-items-absolute-center')
    if assesment:
        assesment = assesment.text[6:9]
    else:
        assesment = None
    start_date = soup.find('div',
                           class_='startdate rc-StartDateString'
                                  ' caption-text').text
    course_duration = soup.find_all('div',
                                    class_='week-heading body-2-text')
    if course_duration:
        course_duration = course_duration[-1].text
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
                'Course duration'])
    for course_info in courses_info:
        exel_row = [course_info['course_name'],
                    course_info['course_lang'],
                    course_info['assesment'],
                    course_info['start_date'],
                    course_info['course_duration']]
        ws1.append(exel_row)
    wb.save(filename=filepath)


if __name__ == '__main__':
    courses_info = []
    dest_filename = 'course_info_book.xlsx'
    for course_url in get_courses_list():
        print(course_url)
        courses_info.append(get_course_info(course_url))
    output_courses_info_to_xlsx(dest_filename, courses_info)
