from dataclasses import dataclass
import datetime

@dataclass(slots=True) # 효율을 위해 선언된 필드만을 강제
class Parcel:
    tracking_number: int | None = -1
    is_resent: bool = False
    row_number: int | None = -1
    delivery_date: datetime.date | None = None

    @classmethod
    def from_excel_row(cls, tracking_number: int = -1, is_resent: bool = False, row_number: int = -1,
                       delivery_date: datetime.date | None = None):
        return cls(tracking_number, is_resent, row_number, delivery_date)