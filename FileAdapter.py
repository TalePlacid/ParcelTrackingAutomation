import os
import shutil
from ParcelSheet import ParcelSheet

class FileAdapter:
    def __init__(self):
        self.path: str = ""

    def duplicate(self, source_path: str)->str:
        self.path = os.path.dirname(source_path) + "\\updated"
        os.makedirs(self.path, exist_ok=True)

        working_path = self.path + "\\updated_" + os.path.basename(source_path)
        shutil.copy2(source_path, working_path)

        return working_path

    def write_log(self, path, loaded_count, fetched_count, updated_count, parcel_sheet: ParcelSheet):
        # 1. 텍스트 파일을 생성한다.
        file_path = self.path + "\\log.txt"
        file = open(file_path, "w", encoding="utf-8")

        # 2. 로그를 작성한다.
        file.write("[처리 사항]\n")
        file.write(f"경로 : {path}\n")
        file.write(f"검색된 수 : {loaded_count}\n")
        file.write(f"배송완료일이 조회된 수 : {fetched_count}\n")
        file.write(f"엑셀파일에 반영된 수 : {updated_count}\n\n")

        file.write("[변경 사항] 행 번호: 송장번호, 재배송여부, 날짜\n")
        i = 0
        while i < parcel_sheet.length:
            parcel = parcel_sheet[i]
            if parcel.delivery_date is not None:
                is_resent = ""
                if parcel.is_resent:
                    is_resent = "(재배송)"
                file.write(f"{parcel.row_number} : {parcel.tracking_number}{is_resent}, {parcel.delivery_date}\n")
            i += 1

        # 3. 텍스트 파일을 저장한다.
        file.close()