import core.EXCEL
import config.config
import time
import os


class target_table(core.EXCEL.Excel):
    def __init__(self, visible=False) -> None:
        os.system(f"start .\\nt2\\re_lite.exe sync \"py:General\" \"{config.config.global_config["Excel"]["target_table_directory"]}\"")
        time.sleep(5)
        super().__init__(visible)
        self.open()

    def open(self) -> core.EXCEL.ExcelWorkBook:
        """open target table, return a workbook"""
        self.workbook = super().open(
            config.config.global_config["Excel"]["target_table"])
        self.target_sheet = self.workbook[1]
        return self.workbook

    def search_date_start(self, date: str) -> int:
        """return the first row of the date starts"""
        try:
            return self.target_sheet.search(date).row
        except:
            return 0

    def search_PB_last(self, pb: str) -> int:
        """return the last row of the date appears"""
        try:
            return self.target_sheet.search(pb, "C:C", core.EXCEL.XL_PREV)
        except:
            return 0

    def search_dn_for_info(self, dn: str) -> dict:
        """return the information of a row based on DN"""
        res = {}
        dn_cell = self.target_sheet.search(dn, "L:L")
        if dn_cell:
            dn_row = dn_cell.row
            for i in range(1, 25):
                res[self.target_sheet[1, i].value] = self.target_sheet[dn_row, i].value
        return res

    def search_row_for_info(self, row_index: int) -> dict:
        res = {}
        for i in range(1, 25):
            res[self.target_sheet[1, i].value] = self.target_sheet[row_index, i].value
        return res

    def check_row_ready_for_report(self, row_index: int):
        try:
            if self.target_sheet[row_index, 12].background_color == 5296274.0:
                return True
            return False
        except:
            return False

    def check_shipping_code(self, dn: str) -> str:
        dn_data = self.search_dn_for_info(dn)
        dest = dn_data["Ship to"]
        if dest.lower() == "sc":
            return "SJ03"
