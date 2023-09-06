import re

import openpyxl

from common import constant_jd


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


text = convert_xlsx_to_text('cn_vn+jp/A604.xlsx')
print(text)


def location(from_word, end_word, text):
    pattern = r'({}) (.+?) ({})'.format(from_word, end_word)
    match = re.search(pattern, text)
    if match:
        return match.group(2)
    else:
        return None


def check_supper(text):
    result = []
    for word in text.split():
        if word.isupper():
            result.append(word)

    return " ".join(result)

def house_money(text):
    if "CÓ" in text:
        return re.sub(r'\s+', ' ',  text.replace("Không", ""))
    else:
        return "KHÔNG"

def extratext(text):
    status = location(constant_jd.JOB_STATUS, constant_jd.NATIONALITY, text)

    nationality = location(constant_jd.NATIONALITY, constant_jd.GENDER, text)

    gender = location(constant_jd.GENDER, constant_jd.CURRENT_LOCATION, text)

    position = location(constant_jd.CURRENT_LOCATION, constant_jd.BRANCH, text)

    branch = location(constant_jd.BRANCH, constant_jd.CAREER, text)

    career = location(constant_jd.CAREER, constant_jd.JAPANESE_ABILITY, text)

    japanese = location(constant_jd.JAPANESE_ABILITY, constant_jd.WORK_LOCATION, text)

    word_location = location(constant_jd.WORK_LOCATION, constant_jd.HIGHLIGHT, text)

    highlight = location(constant_jd.HIGHLIGHT, constant_jd.NUMBER_RECRUITMENT, text)

    number_recruitment = location(constant_jd.NUMBER_RECRUITMENT, constant_jd.WORK_CONTENT, text)

    word_content = location(constant_jd.WORK_CONTENT, constant_jd.EXPERIENCE, text)

    experience = location(constant_jd.EXPERIENCE, constant_jd.WAGE, text)

    wage = location(constant_jd.WAGE, constant_jd.OVERTIME, text)

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
    print('______________________________________________')
    print('BẮT BUỘC')

    print("Tình trạng Job:", status)
    print("Quốc tịch:", nationality)
    print("Giới tính:", gender)
    print("Vị trí hiện tại:", position)
    print("Ngành:", branch)
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

#
extratext(text)
