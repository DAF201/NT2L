import win32com.client
import time
import subprocess
import config.config


class SAPWidget:
    def __init__(self, widget):
        self.widget = widget
        self.type = self.widget.Type
        self.id = self.widget.Id

    def press(self) -> None:
        self.widget.press()

    def enter(self) -> None:
        self.widget.sendVKey(0)

    @property
    def text(self) -> str:
        return self.widget.text

    @text.setter
    def text(self, content: str) -> None:
        self.widget.text = content


class SAPSession:
    id_login_client = "wnd[0]/usr/txtRSYST-MANDT"
    id_login_user = "wnd[0]/usr/txtRSYST-BNAME"
    id_login_password = "wnd[0]/usr/pwdRSYST-BCODE"
    id_login_language = "wnd[0]/usr/txtRSYST-LANGU"
    id_login_comfirm = "wnd[0]/tbar[0]/btn[0]"

    id_login_enter = "wnd[0]/tbar[0]/btn[0]"
    id_command_field = "wnd[0]/tbar[0]/okcd"
    id_main_pannel = "wnd[0]"
    id_backward_button = "wnd[0]/tbar[0]/btn[3]"
    id_save_button = "wnd[0]/tbar[0]/btn[11]"

    id_so_page_command = "/nva01"
    # SO Order Type
    id_so_sales_document_type = "wnd[0]/usr/ctxtVBAK-AUART"
    # SO Sales Organization
    id_so_sales_organization = "wnd[0]/usr/ctxtVBAK-VKORG"
    # SO Distribution Channel
    id_so_distribution_channel = "wnd[0]/usr/ctxtVBAK-VTWEG"
    # SO Division
    id_so_division = "wnd[0]/usr/ctxtVBAK-SPART"
    # SO inner page Sold to Party
    id_so_sold_to_party = "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR"
    # SO inner page Ship to Party
    id_so_ship_to_party = "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR"
    # SO inner page Customer Reference (NV DN)
    id_so_cust_refference = "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD"
    # SO inner page FOX PN
    id_so_material_number = r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]"
    # SO inner page NV PN
    id_so_material_number_used_by_customer = r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-KDMAT[4,0]"
    # SO inner page sale units
    id_so_sale_units = r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,0]"
    # SO innter page amount
    id_so_condition_amount_or_percentage = r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtKOMV-KBETR[9,0]"

    # DN Order (so number text area)
    id_dn_sale_document = "wnd[0]/usr/ctxtLV50C-VBELN"
    # DN shipping and receiving point
    id_dn_shipping_receiving_point = "wnd[0]/usr/ctxtLIKP-VSTEL"
    # DN inner page storage location (SLoc)
    id_dn_storage_location = r"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER/ctxtLIPS-LGORT[3,0]"
    # DN DN value(DN text area on /nva02n)
    id_DN_delivery = "wnd[0]/usr/ctxtLIKP-VBELN"

    # pop out message
    id_pop_out_close = "wnd[1]/tbar[0]/btn[0]"

    def __init__(self, session: win32com.client):
        self.session = session

    def login(self, client, user, password, language) -> None:
        self[SAPSession.id_login_client].text = client
        self[SAPSession.id_login_user].text = user
        self[SAPSession.id_login_password].text = password
        self[SAPSession.id_login_language].text = language
        self[SAPSession.id_login_comfirm].press()

    def __getitem__(self, id) -> "SAPWidget":
        return SAPWidget(self.session.findById(id))

    def get(self, id) -> win32com.client:
        return self.session.findById(id)

    def has_pop_out(self):
        try:
            self.session.FindById("wnd[1]")
            return True
        except:
            return False


class SAP:

    def __init__(self):

        try:
            self.SAP_GUI = win32com.client.GetObject("SAPGUI")
            self.application = self.SAP_GUI.GetScriptingEngine
        except:
            time.sleep(1)
            subprocess.Popen(
                config.config.global_config["SAP"]["path"])
            time.sleep(4)
            self.SAP_GUI = win32com.client.GetObject("SAPGUI")
            self.application = self.SAP_GUI.GetScriptingEngine
        try:
            self.connection = self.application.Children(0)
            self.session = SAPSession(self.connection.Children(0))
        except:
            self.login(
                config.config.global_config["SAP"]["client"], config.config.global_config["SAP"]["user"],
                config.config.global_config["SAP"]["password"], config.config.global_config["SAP"]["language"])

    def login(self, client_id: str, user: str, password: str, language: str) -> None:
        self.connection = self.application.OpenConnection("[W1]WI FPA", True)
        self.session = SAPSession(self.connection.Children(0))
        self.session.login(client_id, user, password, language)
