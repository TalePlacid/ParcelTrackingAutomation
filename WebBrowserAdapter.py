from playwright.async_api import async_playwright, Page, Browser, BrowserContext, Playwright
from dotenv import load_dotenv
import os
import sys
from ParcelSheet import ParcelSheet
from DeliveryDateParser import DeliveryDateParser

# 설정파일 로드
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS        # PyInstaller 임시 폴더
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # 파이참 실행 시 현재 파일 경로
ENV_PATH = os.path.join(BASE_DIR, ".env")
load_dotenv(ENV_PATH)

# 설정파일에서 상수 설정
SITE_ADDRESS = os.getenv("SITE_ADDRESS")
INPUT_FORM = os.getenv("INPUT_FORM")
SUBMIT_BUTTON = os.getenv("SUBMIT_BUTTON")
DEFINITION_LIST = os.getenv("DEFINITION_LIST")
DEFINITION_TERM = os.getenv("DEFINITION_TERM")
DEFINITION_DESCRIPTION = os.getenv("DEFINITION_DESCRIPTION")

class WebBrowserAdapter:
    def __init__(self):
        self.play_wright = None
        self.browser = None
        self.context = None
        self.page = None

    async def fetch(self, parcel_sheet:ParcelSheet)->int:
        # 1. 웹 브라우저룰 시동한다.
        self.play_wright = await async_playwright().start()
        self.browser = await self.play_wright.chromium.launch(channel="chrome", headless=False) # 진행이 보이게 가시화
        self.context = await self.browser.new_context()
        self.page = await self.context.new_page()

        # 2. 배송기록표의 처음부터 끝까지 반복한다.
        count = 0
        i = 0
        while i < parcel_sheet.length:
            # 2.1. 배송기록표에서 10개씩 읽는다.
            tracking_numbers = []
            l = i
            j = 0
            while j < 10 and l < parcel_sheet.length:
                tracking_numbers.append(parcel_sheet[l].tracking_number)
                j += 1
                l += 1

            # 2.2. 배송조회 사이트로 이동한다.
            await self.page.goto(SITE_ADDRESS)

            # 2.3. 입력 폼에 입력한다.
            j = 0
            while j < len(tracking_numbers):
                await self.page.wait_for_selector(f"{INPUT_FORM}{j + 1}", state="visible", timeout=10000) # 로딩 대기
                await self.page.fill(f"{INPUT_FORM}{j + 1}", f"{tracking_numbers[j]:012d}")
                j += 1

            del tracking_numbers

            # 2.4. 조회 버튼을 누른다.
            await self.page.wait_for_selector(SUBMIT_BUTTON, state="visible", timeout=10000)
            await self.page.click(SUBMIT_BUTTON)

            # 2.5. 조회 결과를 스크랩핑한다.
            await self.page.wait_for_selector(DEFINITION_LIST, timeout=15000) # 로딩 대기
            delivery_dates = []
            for j in range(1, 11):
                definition_term = self.page.locator(f"{DEFINITION_LIST} > {DEFINITION_TERM}{j}")
                definition_description = definition_term.locator(DEFINITION_DESCRIPTION)

                if await definition_description.count() > 0:
                    date = (await definition_description.first.inner_text()).strip()
                else:
                    date = None
                delivery_dates.append(date)

                j += 1

            # 2.6 배송기록표에 기재한다.
            l = i
            j = 0
            while j < len(delivery_dates) and l < parcel_sheet.length:
                delivery_date = DeliveryDateParser.parse(delivery_dates[j])
                if delivery_date is not None:
                    parcel_sheet.fill(l, delivery_date)
                    count += 1
                l += 1
                j += 1

            i += j
            del delivery_dates

        # 3. 웹 브라우저를 정지한다.
        if self.context is not None:
            await self.context.close()
        if self.browser is not None:
            await self.browser.close()
        if self.play_wright is not None:
            await self.play_wright.stop()

        # TODO: 예외처리 추가 업데이트

        return count