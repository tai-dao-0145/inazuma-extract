from enum import Enum


class Status(Enum):
    OPEN = "MỞ"
    CLOSE = "ĐÓNG"
    EXPIRED = "HẾT HẠN"


class Nationality(Enum):
    VIETNAM = "Việt Nam"
    INDONESIA = "Indonesia"
    NOT_REQUIRED = "Không yêu cầu"


class Gender(Enum):
    MALE = "NAM"
    FEMALE = "NỮ"
    NOT_REQUIRED = "KHÔNG YÊU CẦU"

class PlaceOfApplication(Enum):
    JAPAN = "Nhật Bản"
    FOREIGN = "Nước ngoài"
    BOTH = "Cả hai"


class JapaneseLevel(Enum):
    N1 = "N1"
    N2 = "N2"
    N3 = "N3"
    N4 = "N4"

