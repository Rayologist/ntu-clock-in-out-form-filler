from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

ROW_START, ROW_END = 2, 33
FIRST_HALF_COL_START, FIRST_HALF_COL_END = 0, 3
SECOND_HALF_COL_START, SECOND_HALF_COL_END = 4, 7


class Table:
    def __init__(self, doc):
        self.table = doc.tables[0]
        self.table.style.font.name = "BiauKai"

    def unitize(self, row_start=ROW_START, row_end=ROW_END):
        first_half = []
        secodn_half = []

        for i in range(row_start, row_end + 1):
            row = self.table.row_cells(i)
            first_half += row[FIRST_HALF_COL_START : FIRST_HALF_COL_END + 1]
            secodn_half += row[SECOND_HALF_COL_START : SECOND_HALF_COL_END + 1]

        all_ = first_half + secodn_half

        indexed_cell = {}

        # A cell unit contains 8 element: first 4 are row1, second 4, row2
        for i in range(len(all_) // 8):
            cell = all_[8 * i : 8 * i + 8]
            indexed_cell[i] = {
                "row1": cell[:4],
                "row2": cell[4:],
            }
        return indexed_cell
    
    def render_dict(self):
        return self.unitize()



class Cell:
    def __init__(self, cell) -> None:
        self._cell = cell

    def apply_basic_format(self, paragraph):
        paragraph.style.font.size = Pt(14)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def add_text(self, text):
        paragraph_in_cell = self._cell.paragraphs[0]
        paragraph_in_cell.text = text
        self.apply_basic_format(paragraph_in_cell)

    def add_signature(self, signature_path):
        paragraph_in_cell = self._cell.paragraphs[0]
        run = paragraph_in_cell.add_run()
        run.add_tab()
        run.add_picture(signature_path, width=Cm(2.8), height=Cm(1.15))

    def __repr__(self):
        return self._cell.paragraphs[0].text

    __str__ = __repr__


class CellUnit:
    def __init__(self, cell_unit):
        self._col1 = cell_unit["row1"]
        self._col2 = cell_unit["row2"]

    @property
    def date(self):
        return Cell(self._col1[0])

    @property
    def work_hours(self):
        return Cell(self._col1[3])

    @property
    def signature_start(self):
        return Cell(self._col1[1])

    @property
    def signature_end(self):
        return Cell(self._col1[2])

    @property
    def start_time(self):
        return Cell(self._col2[1])

    @property
    def end_time(self):
        return Cell(self._col2[2])

    def fill_data(self, **kwargs):
        date = kwargs.get("date")
        start_time = kwargs.get("start_time")
        end_time = kwargs.get("end_time")
        signature_start = kwargs.get("signature_start")
        signature_end = kwargs.get("signature_end")
        work_hours = kwargs.get("work_hours")
        if not all(
            [date, start_time, end_time, signature_end, signature_start, work_hours]
        ):
            raise ValueError(f"None in: {kwargs}")
        self.date.add_text(date)
        self.start_time.add_text(start_time)
        self.end_time.add_text(end_time)
        self.signature_end.add_signature(signature_end)
        self.signature_start.add_signature(signature_start)
        self.work_hours.add_text(work_hours)



class CellData:
    def __init__(self, year, month, start_time, work_hours, work_day, signature_path):
        self._start_year_month = f"{year}/{month}"
        self._end_year_month = f"{year}/{month + 1}"
        self._start_time = pd.Timestamp(start_time)
        self._end_time = pd.Timestamp(start_time) + pd.to_timedelta(work_hours, unit="H")
        self._work_hours = work_hours
        self._work_day = int(work_day)
        self._signature_path = signature_path

    def render_dict(self, method="index"):
        time_table = pd.DataFrame()

        time_table["date"] = (
            pd.date_range(
                start=self._start_year_month,
                end=self._end_year_month,
                closed="left",
                freq="B",
            )
        ).strftime("%m/%d")

        time_table["start_time"] = self._start_time.strftime("%H:%M")
        time_table["end_time"] = self._end_time.strftime("%H:%M")
        time_table["work_hours"] = str(self._work_hours)
        time_table["signature_start"] = self._signature_path
        time_table["signature_end"] = self._signature_path

        time_table = (
            time_table.sample(n=self._work_day)
            .sort_values("date")
            .reset_index(drop=True)
        )

        assert (
            len(time_table) == self._work_day
        ), f"table length ({len(time_table)}) not equal to work day ({self._work_day})"

        return time_table.to_dict(method)