from enum import Enum


class Status(Enum):
    OPEN = "Mở"
    CLOSE = "Đóng"
    EXPIRED = "Hết hạn"


class Nationality(Enum):
    VIETNAM = "Việt Nam"
    INDONESIA = "Indonesia"
    NOT_REQUIRED = "Không yêu cầu"


class Gender(Enum):
    MALE = "Nam"
    FEMALE = "Nữ"
    NOT_REQUIRED = "Không yêu cầu"


class PlaceOfApplication(Enum):
    JAPAN = "Nhật Bản"
    FOREIGN = "Nước ngoài"
    BOTH = "Cả hai"


class JapaneseAbility(Enum):
    N1 = "N1"
    N2 = "N2"
    N3 = "N3"
    N4 = "N4"

