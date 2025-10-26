from ParcelSheet import ParcelSheet
from tkinter import Tk, filedialog, messagebox
from FileAdapter import FileAdapter
from ExcelAdapter import ExcelAdaptor
from WebBrowserAdapter import WebBrowserAdapter
import asyncio

class BatchRunner:
    path: str
    parcelSheet: ParcelSheet | None
    fileAdapter: FileAdapter | None
    excelAdapter: ExcelAdaptor | None

    def __init__(self):
        self.path = ""
        self.parcel_sheet = ParcelSheet()
        self.file_adapter = None
        self.excel_adapter = None
        self.webBrowser_adapter = None

    def run(self):
        # 1. 루트 윈도우를 생성한다.
        root = Tk()  # 루트 윈도우 생성
        root.withdraw()  # 루트 윈도우 숨기기

        # 2. 경로를 선택한다.
        source_path = filedialog.askopenfilename(
            parent = root,
            initialdir = "./",
            title = "파일 선택",
            filetypes = [("Excel 파일", "*.xlsx;*.xls"), ("모든 파일", "*.*")]
        )

        # 3. 경로가 선택되었다면,
        if source_path != "":
            # 3.1. 엑셀 파일을 복사한다.
            self.file_adapter = FileAdapter()
            self.path = self.file_adapter.duplicate(source_path)

            # 3.2. 엑셀파일에서 택배들을 적재한다.
            self.excel_adapter = ExcelAdaptor(self.path)
            loaded_count = self.excel_adapter.load(self.parcel_sheet)

            # 3.3. 배송조회 사이트에서 조회한다.
            self.webBrowser_adapter = WebBrowserAdapter()
            fetched_count = asyncio.run(self.webBrowser_adapter.fetch(self.parcel_sheet))

            # 3.4. 엑셀파일에 반영한다.
            updated_count = self.excel_adapter.update(self.parcel_sheet)

            # 3.5. 로그를 작성한다.
            self.file_adapter.write_log(self.path, loaded_count, fetched_count, updated_count, self.parcel_sheet)

        # 4. 경로가 선택되지 않았다면,
        else:
            # 4.1. 경고 메시지 박스를 띄워준다.
            messagebox.showwarning("경고", "경로가 선택되지 않았습니다.")

if __name__ == "__main__":
    batch_runner = BatchRunner()
    batch_runner.run()