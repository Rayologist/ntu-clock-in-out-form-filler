from .table import Table, CellUnit, CellData
from .header import BasicInfo
from docx import Document
from pathlib import Path


class ClockInOutForm:
    def __init__(self) -> None:
        template_path = Path(__file__).absolute().parent / "templates" / "template.docx"
        self.doc = Document(template_path)

    def add_header_info(
        self,
        year: str,
        month: str,
        department: str,
        name: str,
        expected: str,
        actual: str,
    ) -> None:
        info = BasicInfo(self.doc)
        info.year = str(year)
        info.month = str(month)
        info.department = department
        info.name = name
        info.expected = str(expected)
        info.actual = str(actual)

    def add_cell_data(
        self, year, month, start_time, work_hours, work_day, signature_path
    ) -> None:
        table = Table(self.doc)
        cell_units = table.render_dict()

        cell_data = CellData(
            year=year,
            month=month,
            start_time=start_time,
            work_hours=work_hours,
            work_day=work_day,
            signature_path=signature_path,
        )
        cell_data = cell_data.render_dict()

        for i in range(len(cell_data)):
            cell_unit = CellUnit(cell_units[i])
            cell_unit.fill_data(**cell_data[i])

    def save(self, docx_path: str, open_: bool = False) -> None:
        self.doc.save(docx_path)

        if open_:
            import os, sys

            if sys.platform == "win32":
                os.system(f"start {docx_path}")
            elif sys.platform == "darwin":
                os.system(f"open {docx_path}")
