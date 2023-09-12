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


# print(convert_xlsx_to_text('jd_example/7.xlsx'))

def location(from_word, end_word, text):
    pattern = r'({}) (.+?) ({})'.format(from_word, end_word)
    match = re.search(pattern, text)
    if match:
        return match.group(2)
    else:
        return ''

def extratext_status(text):
    return location(constant_jd.JAPANESE_LEVEL, constant_jd.WORK_LOCATION, text).strip()
def extratext_japanese_level(text):
    return location(constant_jd.JAPANESE_LEVEL, constant_jd.WORK_LOCATION, text).strip()

def extratext_carrer_master(text):
    return location(constant_jd.WORK_INDUSTRY, constant_jd.CAREER, text).strip()


def extratext_word_location(text):
    return location(constant_jd.WORK_LOCATION, constant_jd.HIGHLIGHT, text).strip()


def extratext_highlight(text):
    return location(constant_jd.HIGHLIGHT, constant_jd.NUMBER_RECRUITMENT, text).strip()


def extratext_number_recruitment(text):
    return location(constant_jd.NUMBER_RECRUITMENT, constant_jd.WORK_CONTENT, text).strip()


def extratext_word_content(text):
    return location(constant_jd.WORK_CONTENT, constant_jd.EXPERIENCE, text).strip()


def extratext_experience(text):
    return location(constant_jd.EXPERIENCE, constant_jd.SALARY, text).strip()


def extratext_wage(text):
    return location(constant_jd.SALARY, constant_jd.OVERTIME, text).strip()


def extratext_overtime(text):
    return location(constant_jd.OVERTIME, constant_jd.ESTIMATED_MONTHLY_SALARY, text).strip()


def extratext_estimated_monthly_salary(text):
    return location(constant_jd.ESTIMATED_MONTHLY_SALARY, constant_jd.ESTIMATED_INCOME, text).strip()


def extratext_estimated_income(text):
    return location(constant_jd.ESTIMATED_INCOME, constant_jd.BONUS, text).strip()


def extratext_bonus(text):
    return location(constant_jd.BONUS, constant_jd.SALARY_INCREASE, text).strip()


def extratext_salary_increase(text):
    return location(constant_jd.SALARY_INCREASE, constant_jd.WORKING_TIME, text).strip()


def extratext_working_time(text):
    return location(constant_jd.WORKING_TIME, constant_jd.COMPANY_HOUSE, text).strip()


def extratext_company_house(text):
    return location(constant_jd.COMPANY_HOUSE, constant_jd.SUBSIDIZE, text).strip()


def extratext_subsidize(text):
    return location(constant_jd.SUBSIDIZE, constant_jd.DAY_OFF, text).strip()


def extratext_day_off(text):
    return location(constant_jd.DAY_OFF, constant_jd.WELFARE, text).strip()


def extratext_welfare(text):
    return location(constant_jd.WELFARE, constant_jd.ONBOARD, text).strip()


def extratext_onboard(text):
    return location(constant_jd.ONBOARD, constant_jd.MORE_INFORMATION, text).strip()


def extratext_more_information(text):
    return text.rpartition(constant_jd.MORE_INFORMATION)[-1].strip()


def checkDuplicate(text, array):
    if text not in array:
        array.append(text)


career_master = []
japanese_level = []
word_location = []
highlight = []
number_recruitment = []
word_content = []
experience = []
wage = []
overtime = []

estimated_monthly_salary = []
estimated_income = []
bonus = []
salary_increase = []
working_time = []
company_house = []
subsidize = []
day_off = []
welfare = []
onboard = []
more_information = []

for i in range(2, 23, 1):
    text = convert_xlsx_to_text(r"jd_example/" + str(i) + ".xlsx")

    extracted_career_master = extratext_carrer_master(text)
    checkDuplicate(extracted_career_master, career_master)

    extracted_japanese_level = extratext_japanese_level(text)
    checkDuplicate(extracted_japanese_level, japanese_level)

    extracted_word_location = extratext_word_location(text)
    checkDuplicate(extracted_word_location, word_location)

    extracted_highlight = extratext_highlight(text)
    checkDuplicate(extracted_highlight, highlight)

    extracted_number_recruitment = extratext_number_recruitment(text)
    checkDuplicate(extracted_number_recruitment, number_recruitment)

    extracted_word_content = extratext_word_content(text)
    checkDuplicate(extracted_word_content, word_content)

    extracted_experience = extratext_experience(text)
    checkDuplicate(extracted_experience, experience)

    extracted_wage = extratext_wage(text)
    checkDuplicate(extracted_wage, wage)

    extracted_overtime = extratext_overtime(text)
    checkDuplicate(extracted_overtime, overtime)

    extracted_estimated_monthly_salary = extratext_estimated_monthly_salary(text)
    checkDuplicate(extracted_estimated_monthly_salary, estimated_monthly_salary)

    extracted_estimated_income = extratext_estimated_income(text)
    checkDuplicate(extracted_estimated_income, estimated_income)

    extracted_bonus = extratext_bonus(text)
    checkDuplicate(extracted_bonus, bonus)

    extracted_salary_increase = extratext_salary_increase(text)
    checkDuplicate(extracted_salary_increase, salary_increase)

    extracted_working_time = extratext_working_time(text)
    checkDuplicate(extracted_working_time, working_time)

    extracted_company_house = extratext_company_house(text)
    checkDuplicate(extracted_company_house, company_house)

    extracted_subsidize = extratext_subsidize(text)
    checkDuplicate(extracted_subsidize, subsidize)

    extracted_day_off = extratext_day_off(text)
    checkDuplicate(extracted_day_off, day_off)

    extracted_welfare = extratext_welfare(text)
    checkDuplicate(extracted_welfare, welfare)

    extracted_onboard = extratext_onboard(text)
    checkDuplicate(extracted_onboard, onboard)

print(career_master)
print(japanese_level)
print(word_location)
print(highlight)
print(number_recruitment)
print(word_content)
print(experience)
print(wage)
print(overtime)

print(estimated_monthly_salary)
print(estimated_income)
print(bonus)
print(salary_increase)
print(working_time)
print(company_house)
print(subsidize)
print(day_off)
print(welfare)
print(onboard)
print(more_information)


def write_arrays_to_txt(content, output_file):
    try:
        with open(output_file, 'w', encoding='utf-8') as file:
            for item in content:
                file.write(item + '\n')

    except Exception as e:
        print(f"An error occurred: {str(e)}")

output_japanese_level = "tiengnhat.txt"
output_career_master = "nghanh.txt"
output_word_location = "diadiemlamviec.txt"
output_highlight = "diemnoibat.txt"
output_number_recruitment = "soluongtuyendung.txt"
output_word_content = "noidungcongviec.txt"
output_experience = "kinhnghiem.txt"
output_wage = "luong.txt"
output_overtime = "tangca.txt"
output_estimated_monthly_salary = "mucluonghangthang.txt"
output_estimated_income = "uoctinhvetay.txt"
output_bonus = "thuong.txt"
output_salary_increase = "tangluong.txt"
output_working_time = "thoigianlamviec.txt"
output_company_house = "nhaocongty.txt"
output_subsidize = "trocap.txt"
output_day_off = "ngaynghi.txt"
output_welfare = "phucloi.txt"
output_onboard = "thoidiemvaocty.txt"

# write_arrays_to_txt(japanese_level, output_japanese_level)
# write_arrays_to_txt(career_master, output_career_master)
# write_arrays_to_txt(word_location, output_word_location)
# write_arrays_to_txt(highlight, output_highlight)
# write_arrays_to_txt(number_recruitment, output_number_recruitment)
# write_arrays_to_txt(word_content, output_word_content)
# write_arrays_to_txt(experience, output_experience)
# write_arrays_to_txt(wage, output_wage)
# write_arrays_to_txt(overtime, output_overtime)
# write_arrays_to_txt(estimated_monthly_salary, output_estimated_monthly_salary)
# write_arrays_to_txt(estimated_income, output_estimated_income)
# write_arrays_to_txt(bonus, output_bonus)
# write_arrays_to_txt(salary_increase, output_salary_increase)
# write_arrays_to_txt(working_time, output_working_time)
# write_arrays_to_txt(company_house, output_company_house)
# write_arrays_to_txt(subsidize, output_subsidize)
# write_arrays_to_txt(day_off, output_day_off)
# write_arrays_to_txt(welfare, output_welfare)
# write_arrays_to_txt(onboard, output_onboard)
