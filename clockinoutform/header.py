from ._types import Document, Run


class RunText:
    """
    The RunText object is a wrapper class for docx.text.run.Run, adding space padding while modifying the text.
    """

    def __init__(self, run: Run) -> None:
        if not isinstance(run, Run):
            raise TypeError(f"{run}: {type(run).__name__} not an instance of Run")
        self._run = run

    def add_text(self, text: str, padding_length: int = 2) -> str:
        self._run.text = self._pad(text, padding_length, " ")
        return self._run.text

    def _pad(self, text: str, length: int, padding: str) -> str:
        length_rjust = len(text) + length
        length_ljust = len(text) + length * 2
        return text.rjust(length_rjust, padding).ljust(length_ljust, padding)


class HeaderInfo:
    """
    The HeaderInfo object handles adding/modifying texts to the header in the template.docx.
    """

    def __init__(self, doc: Document) -> None:
        self.doc = doc

    @property
    def year(self) -> Run:
        addr = self.doc.paragraphs[0].runs[1]
        return addr

    @year.setter
    def year(self, value: str) -> None:
        run = RunText(self.year[0])
        run.add_text(value)

    @property
    def month(self) -> Run:
        addr = self.doc.paragraphs[0].runs[3]
        return addr

    @month.setter
    def month(self, value: str) -> None:
        value = str(value).zfill(2)
        run = RunText(self.month[0])
        run.add_text(value)

    @property
    def department(self) -> Run:
        addr = self.doc.paragraphs[1].runs[1]
        return addr

    @department.setter
    def department(self, value: str) -> None:
        run = RunText(self.department[0])
        run.add_text(value)

    @property
    def name(self) -> Run:
        addr = self.doc.paragraphs[1].runs[6]
        return addr

    @name.setter
    def name(self, value: str) -> None:
        run = RunText(self.name[0])
        run.add_text(value)

    @property
    def expected(self) -> Run:
        addr = self.doc.paragraphs[1].runs[9]
        return addr

    @expected.setter
    def expected(self, value: str) -> None:
        run = RunText(self.expected[0])
        run.add_text(value)

    @property
    def actual(self) -> Run:
        addr = self.doc.paragraphs[1].runs[13]
        return addr

    @actual.setter
    def actual(self, value: str) -> None:
        run = RunText(self.actual[0])
        run.add_text(value)
