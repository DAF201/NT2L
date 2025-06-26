import os
import win32com.client as win32
import time
import shutil

# Microsoft's constants for search and others operations
XL_PASTE_ALL = -4104
XL_PASTE_VALUES = -4163
XL_PASTE_FORMATS = -4122
XL_VALUES = -4163
XL_WHOLE = 1
XL_PART = 2
XL_BY_ROWS = 1
XL_BY_COLUMNS = 2
XL_NEXT = 1
XL_PREV = 2


class ExcelCell:
    """excel cell can be operated on"""

    def __init__(self, cell: win32) -> None:
        self.cell = cell

    @property
    def value(self) -> object:
        return self.cell.Value

    @value.setter
    def value(self, val: object) -> None:
        self.cell.Value = val

    @property
    def row(self) -> int:
        return self.cell.Row

    @property
    def col(self) -> int:
        return self.cell.Column

    @property
    def background_color(self) -> int:
        return self.cell.Interior.Color

    @background_color.setter
    def background_color(self, color: int) -> None:
        self.cell.Interior.Color = color

    @property
    def text_color(self) -> int:
        return self.cell.Font.Color

    @text_color.setter
    def text_color(self, color: int) -> None:
        self.cell.Font.Color = color

    def clear(self) -> None:
        """clear the contents in this cell"""
        self.cell.Clear()

    def copy(self) -> None:
        """copy the contents in this cell"""
        self.cell.Copy()

    def paste_special(self, paste_type=XL_PASTE_VALUES) -> None:
        """paste copied contents to this cell"""
        self.cell.PasteSpecial(Paste=paste_type)

    def copy_to(self, dst_cell: 'ExcelCell', copy_format=True):
        """copy value, colors to another cell (may in different sheet or workbook)"""
        dst_cell.value = self.value
        if copy_format:
            dst_cell.background_color = self.background_color
            dst_cell.text_color = self.text_color

    def select(self):
        """select this cell"""
        self.cell.Select()


class ExcelSheet:
    """a excel work sheet can be operated on"""

    def __init__(self, sheet: win32) -> None:
        self.sheet = sheet

    def __getitem__(self, index: int) -> ExcelCell | None:
        """use sheet[1,1] or such things to get a excel cell object can be worked on"""
        if isinstance(index, tuple) and len(index) == 2:
            row, col = index
            if isinstance(row, int) and isinstance(col, int):
                return ExcelCell(self.sheet.Cells(row, col))
            else:
                raise TypeError("Row and column indices must be integers")
        else:
            raise IndexError(
                "Index must be a tuple of two integers (row, column)")

    def clear(self, range=None) -> None:
        """remove everything on this sheet"""
        if range == None:
            self.sheet.Cells.Clear()
        else:
            self.sheet.Range(range).Clear()

    def range(self, address: str) -> object:
        """select a range on this sheet, 'A1:B5' or so"""
        return self.sheet.Range(address)

    def search(self, key: str, search_range="", search_direction=XL_NEXT) -> ExcelCell | None:
        """search for a value in this sheet/range, will return the first cell found"""
        target_range = self.sheet.UsedRange if not search_range else self.range(
            search_range)

        found = target_range.Find(
            What=key,
            LookIn=XL_VALUES,
            LookAt=XL_WHOLE,
            SearchOrder=XL_BY_ROWS,
            SearchDirection=search_direction
        )
        return ExcelCell(found) if found else None

    def copy_and_paste(self, src_range: str, dst_sheet: "ExcelSheet", dst: str, paste_type=XL_PASTE_ALL) -> None:
        """copy all the contents in a range to another sheet"""
        self.range(src_range).Copy()
        dst_sheet.sheet.Activate()
        dst_range = dst_sheet.range(dst)
        dst_range.Select()
        time.sleep(0.1)
        dst_range.PasteSpecial(Paste=paste_type)

    def show_all_data(self):
        """remove filters on this sheet"""
        try:
            self.sheet.ShowAllData()
        except:
            pass

class ExcelWorkBook:
    """a workbook COM that can be operated on"""

    def __init__(self, workbook) -> None:
        self.workbook = workbook

    def __getitem__(self, key) -> ExcelSheet:
        """use workbook[1], workbook[2]... to get cordinated sheet(sheet name works too but index is eaiser)"""
        if isinstance(key, int):
            return ExcelSheet(self.workbook.Sheets(key))
        elif isinstance(key, str):
            return ExcelSheet(self.workbook.Sheets(key))
        raise TypeError("Index must be int or str")

    @property
    def sheets(self) -> win32:
        """internal use"""
        return self.workbook.Sheets

    @property
    def sheets_count(self) -> int:
        """get the number of sheets in this workbook"""
        return self.sheets.Count

    def add_sheet(self, index: int, direction="AFTER", name=None) -> None:
        """add a sheet BEFORE or AFTER the index sheet"""
        if name is None:
            name = f"Sheet {index+1}"
        if direction == "BEFORE":
            self.workbook.Sheets.Add(
                Before=self.workbook.Sheets(index)).Name = name
        else:
            self.workbook.Sheets.Add(
                After=self.workbook.Sheets(index)).Name = name

    def move_sheet(self, src_index, dst_index, direction="AFTER") -> None:
        """move a sheet to before or after another sheet in this workbook"""
        sheet = self.sheets(src_index)
        if direction == "BEFORE":
            sheet.Move(Before=self.sheets(dst_index))
        else:
            sheet.Move(After=self.sheets(dst_index))

    def save(self) -> None:
        """save workbook"""
        self.workbook.Save()

    def close(self, save: bool) -> None:
        """let the Excel close workbook, do not explicately call this"""
        self.workbook.Close(SaveChanges=save)


class Excel:
    """Excel COM application, can be opened to get a workbookk COM"""

    def __init__(self, visible: bool = False) -> None:
        self.app = win32.DispatchEx("Excel.Application")
        self.app.Visible = visible
        self.workbook = None
        self.path = None
        self.closed = True

    def open(self, file_path: str) -> ExcelWorkBook:
        """assign a xlsx file to open, then return a wrapped workbook object"""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        if self.workbook:
            self.close()
        self.workbook = ExcelWorkBook(self.app.Workbooks.Open(file_path))
        self.path = file_path
        self.closed = False
        return self.workbook

    def rename(self, target_path: str) -> None:
        """rename .xlsx file"""
        if not self.closed:
            print("Excel can only be renamed when closed")
            return
        if self.path is not None:
            shutil.move(self.path, target_path)

    def close(self, save: bool = False) -> None:
        """close the associated workbook"""
        if self.workbook:
            self.workbook.close(save)
            self.workbook = None
        self.closed = True

    def quit(self) -> None:
        """quit Excel APP, call this function not close"""
        self.close()
        if self.app:
            self.app.Quit()
            self.app = None
        self.closed = True

    def __enter__(self) -> "Excel":
        """with Excel() as excel: wb=excel.open(path)..."""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """for auto clean up"""
        self.quit()


def col_num_to_letter(n: int) -> str:
    """Convert 1-based column index to Excel-style column letter"""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result
