from dataclasses import dataclass, field
import datetime
from Parcel import Parcel

@dataclass(slots=True)
class ParcelSheet:
    parcels : list[Parcel] = field(default_factory=list)
    length : int = 0

    def record_parcel(self, parcel:Parcel)->int:
        self.parcels.append(parcel)
        self.length += 1
        return self.length - 1

    def record_from_excel_row(self, tracking_number: int, is_resent: bool, row_number: int,
                          delivery_date: datetime.date | None = None)->int:
        parcel = Parcel.from_excel_row(tracking_number, is_resent, row_number, delivery_date)
        self.parcels.append(parcel)
        self.length += 1
        return self.length - 1

    def fill(self, index: int, delivery_date: datetime.date)->int:
        parcel = self.parcels[index]
        corrected_parcel = Parcel.from_excel_row(parcel.tracking_number, parcel.is_resent, parcel.row_number, delivery_date)
        self.parcels[index] = corrected_parcel
        return index

    def __getitem__(self, index: int):
        return self.parcels[index]

#if __name__ == "__main__":
    # 1. 택배기록표를 만든다.
    #parcelSheet = ParcelSheet()

    # 2. 기재한다.
    #index = parcelSheet.record_from_excel_row(11112222333344, False, 1, datetime.date.today())
    #parcel = parcelSheet[index]
    #print(f"{parcel.row_number}: {parcel.tracking_number}, {parcel.is_resent}, {parcel.delivery_date}")

    #index = parcelSheet.record_from_excel_row(55556666777788, True, 2, datetime.date(2025, 9, 1))
    #parcel = parcelSheet[index]
    #print(f"{parcel.row_number}: {parcel.tracking_number}, {parcel.is_resent}, {parcel.delivery_date}")

    #index = parcelSheet.record_parcel(Parcel.from_excel_row(10000000000000, False, 3))
    #parcel = parcelSheet[index]
    #print(f"{parcel.row_number}: {parcel.tracking_number}, {parcel.is_resent}, {parcel.delivery_date}")

    # 3. 채운다.
    #index = parcelSheet.fill(2, datetime.date.today())
    #parcel = parcelSheet[index]
    #print(f"{parcel.row_number}: {parcel.tracking_number}, {parcel.is_resent}, {parcel.delivery_date}")
