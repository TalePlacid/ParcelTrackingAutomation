import re
import datetime
from dotenv import load_dotenv
import os
import sys

# 설정파일 로드
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS        # PyInstaller 임시 폴더
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # 파이참 실행 시 현재 파일 경로
ENV_PATH = os.path.join(BASE_DIR, ".env")
load_dotenv(ENV_PATH)

# 설정파일에서 상수 설정
PATTERN = os.getenv("PATTERN")

class DeliveryDateParser:
    @classmethod
    def parse(cls, delivery_date)->datetime.date | None:
        parsed = None

        # 1. 패턴과 일치하는 지 확인한다.
        if delivery_date is not None:
            pattern = re.compile(PATTERN)
            valid_date = pattern.fullmatch(delivery_date)

            # 2. 패턴과 일치한다면,
            if valid_date:
                # 2.1. 월과 일을 정수로 추출한다.
                mm = int(valid_date.group('m'))
                dd = int(valid_date.group('d'))

                today = datetime.date.today()
                year = today.year

                # 2.2. 조회 시점이 1월이고, 배송완료일이 12월이면, 작년으로 보정한다.
                if  mm == 12 and today.month == 1:
                    year -= 1

                parsed = datetime.date(year, mm, dd)

        return parsed