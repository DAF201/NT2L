import bs4
import requests
import config.config
# mostly copied from old code, I dont remember what exactly each parts do, but I remember our website return value was messy and I dont want to look into it


class SFC:
    def __init__(self, host: str) -> None:
        self.host = host

    def mo_query(self, model_name: str) -> list[dict] | None:
        """SFC mo_query, will return information of a model like working order number, start sn, number of units..."""
        try:
            form = {
                "modelName": model_name,
                "modelNameSubmit": "Query"
            }
            url = f"http://{self.host}/SFCWeb/nvsfc/qa/MO/mo_number_query.asp"
            table = bs4.BeautifulSoup(requests.post(
                url=url, data=form).content.decode("big5"), "html.parser").find("table")

            headers = [th.get_text(strip=True)
                       for th in table.find_all("td", class_="tableheader")]
            data = []
            for tr in table.find_all("tr")[1:]:
                cells = tr.find_all("td")
                if len(cells) != len(headers):
                    continue
                row_dict = {}
                for header, cell in zip(headers, cells):
                    text = "".join(cell.stripped_strings)
                    row_dict[header] = text
                data.append(row_dict)
            return data
        except:
            return None

    def wo_query(self, work_order: str, department="") -> list[dict] | None:
        """get information about the working order"""
        try:
            url = f"http://{self.host}/SFCWeb/nvsfc/workorder/numresult.asp?only=b&gnew={department}&monum={work_order}&model=&line=&group=&st=&ed=&flag=0"
            if department == "":
                url = f"http://{self.host}/SFCWeb/nvsfc/workorder/numresult.asp?only=b&monum={work_order}&model=&line=&group=&st=&ed=&flag=0"

            rows = bs4.BeautifulSoup(requests.get(
                url).content.decode("big5", errors="ignore"), "html.parser").find_all("tr")
            headers = [
                "ID", "Serial Number", "Mo Number", "Model Name", "Version Code",
                "Line Name", "Group Name", "Error Flag", "In Station Time",
                "Container NO", "Carton NO", "Emp Name"
            ]
            data = []
            for row in rows:
                cols = row.find_all("td")
                values = [col.get_text(strip=True).replace(
                    "\xa0", "") for col in cols]
                if len(values) == len(headers):
                    row_dict = dict(zip(headers, values))
                    data.append(row_dict)
            return data
        except:
            return None

    def packing_tracking(self, work_order: str) -> list[str] | None:
        """to find all units in packing, then get the serial number as a list"""
        try:
            serial_numbers = set()
            for line in self.wo_query(work_order, "PACKING"):
                serial_numbers.add(line["Serial Number"])
            return sorted(list(serial_numbers))
        except:
            return None

    def oqc_tracking(self, work_order: str) -> list[str] | None:
        """to find all units in OQC, then get the cartoon id as a list"""
        try:
            cartoon_ids = set()
            for line in self.wo_query(work_order, "OQC"):
                cartoon_ids.add(line["Carton NO"])
            return sorted(list(cartoon_ids))
        except:
            return None

    def sn_lookup(self, serial_number: int | str) -> dict | None:
        """look up infomation about a serial number"""
        try:
            product_tracking_form = {
                "T_SN": str(serial_number)
            }
            url = f"http://{self.host}/SFCWeb/nvsfc/public/product_track.asp"
            rows = bs4.BeautifulSoup(requests.post(
                url, data=product_tracking_form).content, "html.parser").find_all("tr")
            data = []
            for row in rows:
                cols = row.find_all("td")
                data.append([col.get_text(strip=True).replace("\xa0", "")
                            for col in cols])
            data = data[1:]
            res = {"NPI OUT": False}
            i = 0
            for x in data:
                for i in range(0, len(x), 2):
                    if i+1 < len(x):
                        if "NPI_OUT" in x:
                            res["NPI OUT"] = True
                        if "OQC" in x:
                            res["cartoon_id"] = x[8]
                        res[x[i]] = x[i+1]
            return res
        except:
            return None


global_SFC_API = SFC(config.config.global_config["SFC"]["host"])
