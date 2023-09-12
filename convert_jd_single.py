import re

import openpyxl

from common import constant_jd, enum_jd


def convert_xlsx_to_text(file_path):
    workbook = openpyxl.load_workbook(file_path)
    worksheet = workbook.active
    text = ""
    for row in worksheet.iter_rows():
        for cell in row:
            if cell.value:
                if cell.font.bold:
                    text += str(cell.value).upper() + " "
                else:
                    text += str(cell.value) + " "
    return text.strip().replace("\n", " ")


text = convert_xlsx_to_text('jd_example/15.xlsx')
print(text)


def location(from_word, end_word, text):
    pattern = r'({}) (.+?) ({})'.format(from_word, end_word)
    match = re.search(pattern, text)
    if match:
        return match.group(2)
    else:
        return ''

def extratext_full(text):
    code = location(constant_jd.CODE, constant_jd.END_DATE, text)

    end_date = location(constant_jd.END_DATE, constant_jd.STATUS, text)

    status = location(constant_jd.STATUS, constant_jd.NATIONALITY, text)

    nationality = location(constant_jd.NATIONALITY, constant_jd.SEX, text)

    sex = location(constant_jd.SEX, constant_jd.LOCATION_APPLICATION, text)

    position = location(constant_jd.LOCATION_APPLICATION, constant_jd.WORK_INDUSTRY, text)

    work_industry = location(constant_jd.WORK_INDUSTRY, constant_jd.CAREER, text)

    career = location(constant_jd.CAREER, constant_jd.JAPANESE_LEVEL, text)

    japanese = location(constant_jd.JAPANESE_LEVEL, constant_jd.WORK_LOCATION, text)

    word_location = location(constant_jd.WORK_LOCATION, constant_jd.HIGHLIGHT, text)

    highlight = location(constant_jd.HIGHLIGHT, constant_jd.NUMBER_RECRUITMENT, text)

    number_recruitment = location(constant_jd.NUMBER_RECRUITMENT, constant_jd.WORK_CONTENT, text)

    word_content = location(constant_jd.WORK_CONTENT, constant_jd.EXPERIENCE, text)

    experience = location(constant_jd.EXPERIENCE, constant_jd.SALARY, text)

    wage = location(constant_jd.SALARY, constant_jd.OVERTIME, text)

    overtime = location(constant_jd.OVERTIME, constant_jd.ESTIMATED_MONTHLY_SALARY, text)

    estimated_monthly_salary = location(constant_jd.ESTIMATED_MONTHLY_SALARY, constant_jd.ESTIMATED_INCOME, text)

    estimated_income = location(constant_jd.ESTIMATED_INCOME, constant_jd.BONUS, text)

    bonus = location(constant_jd.BONUS, constant_jd.SALARY_INCREASE, text)

    salary_increase = location(constant_jd.SALARY_INCREASE, constant_jd.WORKING_TIME, text)

    working_time = location(constant_jd.WORKING_TIME, constant_jd.COMPANY_HOUSE, text)

    company_house = location(constant_jd.COMPANY_HOUSE, constant_jd.SUBSIDIZE, text)

    subsidize = location(constant_jd.SUBSIDIZE, constant_jd.DAY_OFF, text)

    day_off = location(constant_jd.DAY_OFF, constant_jd.WELFARE, text)

    welfare = location(constant_jd.WELFARE, constant_jd.ONBOARD, text)

    onboard = location(constant_jd.ONBOARD, constant_jd.MORE_INFORMATION, text)

    more_information = text.rpartition(constant_jd.MORE_INFORMATION)[-1]


    print('\n')
    print('Mã đăng tuyển:', code)
    print('Ngày kết thúc tuyển dụng:', end_date)
    print('______________________________________________')
    print('BẮT BUỘC')

    print("Tình trạng Job:", status)
    print("Quốc tịch:", nationality)
    print("Giới tính:", sex)
    print("Vị trí hiện tại:", position)
    print("Ngành:", work_industry)
    print("Nghề:", career)
    print("Năng lực tiếng nhật:", japanese)

    print('______________________________________________')
    print("Địa điểm làm việc:", word_location)
    print("Điểm nổi bật:", highlight)
    print("Số lượng tuyển dụng:", number_recruitment)
    print("Nội dung công việc:", word_content)
    print("Kinh nghiệm:", experience)
    print("Lương:", wage)
    print("Tăng ca:", overtime)
    print("Mức lương ước tính hàng tháng:", estimated_monthly_salary)
    print("Thu nhập về tay ước tính:", estimated_income)
    print("Thưởng:", bonus)
    print("Tăng lương:", salary_increase)
    print("Thời gian làm việc:", working_time)
    print("Nhà ở công ty:", company_house)
    print("Trợ cấp:", subsidize)
    print("Ngày nghỉ:", day_off)
    print("Phúc lợi:", welfare)
    print("Thời điểm vào công ty:", onboard)
    print("Thông tin thêm :", more_information)


extratext_full(text)

def extratext_field_code(text):
    return location(constant_jd.CODE, constant_jd.END_DATE, text).strip()

def extratext_field_end_date(text):
    return location(constant_jd.END_DATE, constant_jd.STATUS, text).strip()

def extratext_field_status_job(text):
    return location(constant_jd.STATUS, constant_jd.NATIONALITY, text).strip()

def extratext_field_nationality(text):
    return location(constant_jd.NATIONALITY, constant_jd.SEX, text).strip()

def extratext_field_sex(text):
    return location(constant_jd.SEX, constant_jd.LOCATION_APPLICATION, text).strip()

def extratext_field_location_application(text):
    return location(constant_jd.LOCATION_APPLICATION, constant_jd.WORK_INDUSTRY, text).strip()

def extratext_field_location_work(text):
    return location(constant_jd.WORK_LOCATION, constant_jd.HIGHLIGHT, text).strip()

def extratext_field_salary(text):
    return location(constant_jd.SALARY, constant_jd.OVERTIME, text).strip()

def extratext_field_overtime(text):
    return location(constant_jd.OVERTIME, constant_jd.ESTIMATED_MONTHLY_SALARY, text).strip()


# print(extratext_field_code(text))
# print(extratext_field_end_date(text))
# print(extratext_field_status_job(text))
# print(extratext_field_nationality(text))
# print(extratext_field_sex(text))
# print(extratext_field_location_application(text))
# print(extratext_field_location_work(text))

def map_sex_to_value(sex):
    if sex.upper() == enum_jd.Gender.MALE.value:
        return 0
    elif sex.upper() == enum_jd.Gender.FEMALE.value:
        return 1
    elif sex.upper() == enum_jd.Gender.NOT_REQUIRED.value:
        return 2
    else:
        return None

def map_status_to_value(status):
    if status.upper() == enum_jd.Status.OPEN.value:
        return 0
    elif status.upper() == enum_jd.Status.CLOSE.value:
        return 1
    elif status.upper() == enum_jd.Status.EXPIRED.value:
        return 2
    else:
        return None


def extratext_hour_salary(salary):
    hourly_salary_pattern = r"hàng giờ\] ([\d.,]+) [Yy][eê]n"
    match = re.search(hourly_salary_pattern, salary)

    if match:
        hourly_salary = match.group(1)
        return hourly_salary
    else:
        return None


def extratext_month_salary(salary):
    monthly_salary_pattern = r"hàng tháng\] ([\d.,]+) [Yy][eê]n"
    match = re.search(monthly_salary_pattern, salary)

    if match:
        monthly_salary = match.group(1)
        return monthly_salary
    else:
        return None



def extratext_overtime_hours(overtime):

    overtime_hours_pattern = r"(\d+)\s*(?:[gh]|giờ|h)\b"
    match = re.search(overtime_hours_pattern, overtime)

    if match:
        overtime_hours = match.group(1)
        return overtime_hours
    else:
        return None


def extract_japanese_level(japanese):
    japanese_level_pattern = r"[Nn]\d"
    match = re.search(japanese_level_pattern, japanese)

    if match:
        japanese_level = match.group()
        return japanese_level
    else:
        return japanese


