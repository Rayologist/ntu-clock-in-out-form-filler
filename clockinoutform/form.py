from .table import TableStacker, Grid, CellDataGenerator
from .header import HeaderInfo
from docx import Document
from pathlib import Path
from ._types import str_or_int


class ClockInOutForm:
    """
    The ClockInOutForm deals with how the text should be added/modified in the template.docx.
    """

    def __init__(self) -> None:
        template_path = Path(__file__).absolute().parent / "templates" / "template.docx"
        self.doc = Document(template_path)

    def fill_header(
        self,
        year: str_or_int,
        month: str_or_int,
        department: str,
        name: str,
        expected: str,
        actual: str,
    ) -> None:
        info = HeaderInfo(self.doc)
        info.year = str(year)
        info.month = str(month)
        info.department = department
        info.name = name
        info.expected = str(expected)
        info.actual = str(actual)

    def fill_table(
        self,
        year: int,
        month: int,
        start_time: str,
        work_hours: int,
        work_day: int,
        signature_path: str,
    ) -> None:

        table = TableStacker(self.doc)
        grids = table.render_dict()

        cell_data = CellDataGenerator(
            year=year,
            month=month,
            start_time=start_time,
            work_hours=work_hours,
            work_day=work_day,
            signature_path=signature_path,
        )
        cell_data = cell_data.render_dict()

        for i in range(len(cell_data)):
            grid = Grid(grids[i])
            grid.fill_data(**cell_data[i])

    def save(self, docx_path: str, open_: bool = False) -> None:

        self.doc.save(docx_path)

        if open_:

            import os, sys

            if sys.platform == "win32":
                os.system(f"start {docx_path}")

            elif sys.platform == "darwin":
                os.system(f"open {docx_path}")
