import re

import openpyxl

from common import constant_cv


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
    return text.strip().replace(constant_cv.CV, "")


text = convert_xlsx_to_text('cv_vn/Phiếu nguyện vọng KYUJIN.xlsx')
print(text)


def location(from_word, end_word, text):
    pattern = r'({}) (.+?) ({})'.format(from_word, end_word)
    match = re.search(pattern, text)
    if match:
        return match.group(2)
    else:
        return None

def extratext(text):

    status = location(constant_cv.CANDIDATE_STATUS, constant_cv.NATIONALITY, text)

    nationality = location(constant_cv.NATIONALITY, constant_cv.GENDER, text)

    gender = location(constant_cv.GENDER, constant_cv.BRANCH, text)

    branch = location(constant_cv.BRANCH, constant_cv.CAREER, text)

    career = location(constant_cv.CAREER, constant_cv.CURRENT_LOCATION, text)

    position = location(constant_cv.CURRENT_LOCATION, constant_cv.JAPANESE_ABILITY, text)

    japanese = location(constant_cv.JAPANESE_ABILITY, constant_cv.EXPERIENCE, text)

    experience = location(constant_cv.EXPERIENCE, constant_cv.ONBOARD, text)

    onboard = location(constant_cv.ONBOARD, constant_cv.WISH, text)

    wish = location(constant_cv.WISH, constant_cv.WORK_LOCATION, text)

    word_location = location(constant_cv.WORK_LOCATION, constant_cv.WAGE, text)

    wage = location(constant_cv.WAGE, constant_cv.OVERTIME, text)

    overtime = location(constant_cv.OVERTIME, constant_cv.WORK_CONTENT, text)

    work_content = location(constant_cv.WORK_CONTENT, constant_cv.WORKING_TIME, text)

    working_time = location(constant_cv.WORKING_TIME, constant_cv.COMPANY_HOUSE, text)

    company_house = location(constant_cv.COMPANY_HOUSE, constant_cv.OTHER_MODES, text)

    other_modes = location(constant_cv.OTHER_MODES, constant_cv.OTHER_WISHES, text)

    other_wishes = text.rpartition(constant_cv.OTHER_WISHES)[-1]


    print('\n')
    print('______________________________________________')


    print("Tình trạng:", status)
    print("Quốc tịch:", nationality)
    print("Giới tính:", gender)
    print("Ngành :", branch)
    print("Nghề :", career)
    print("Vị trí hiện tại:", position)
    print("Chứng chỉ tiếng Nhật:", japanese)

    print("Kinh ngiệm:", experience)
    print("Thời điểm vào công ty:", onboard)
    print("Nguyện vọng:", wish)
    print("Địa điểm làm việc:", word_location)
    print("Lương thưởng:", wage)
    print("Tăng ca:", overtime)
    print("Nội dung công việc:", work_content)
    print("Thời gian làm việc:", working_time)
    print("Nhà ở công ty:", company_house)
    print("Chế độ đãi ngộ khác:", other_modes)
    print("Nguyện vọng khác:", other_wishes)

extratext(text)
