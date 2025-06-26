import uiautomation
import config
import config.config
import subprocess
import time

uiautomation.uiautomation.SetGlobalSearchTimeout(0)


class SAP_Shipping:
    def __init__(self):
        self.path = config.config.global_config["SAP Shipping"]["path"]
        self.start()

    def start(self) -> None:
        """start the SAP shipping.exe if not started"""
        self.main_win = uiautomation.WindowControl(
            Name="Shipping  Workstation ")
        if self.main_win.Exists(0):
            return
        else:
            self.process = subprocess.Popen(self.path)
            time.sleep(3)
            alert_window = uiautomation.WindowControl(Name="Shipping")
            close_button = alert_window.ButtonControl(Name="OK")
            close_button.Click(simulateMove=False)
            self.main_win = uiautomation.WindowControl(
                Name="Shipping  Workstation ")

    def stop(self) -> None:
        """terminate process"""
        self.process.terminate()

    def operator_id(self, value: str) -> None:
        """set operator id, the first thing should be set once entered"""
        operator_id_box = self.main_win.PaneControl(Name="Panel5").PaneControl(Name="Panel6").PaneControl(
            Name="Panel8").PaneControl(Name="OPERATOR").ComboBoxControl(Name="").EditControl(AutomationId="1001")
        operator_id_box.SetFocus()
        operator_id_box.SendKeys(value, waitTime=0)
        operator_id_box.SendKeys("{Enter}", waitTime=0)

    def order(self, value: str) -> None:
        """set order (FXDN)"""
        order_box = self.main_win.PaneControl(Name="").PaneControl(Name="Panel14").PaneControl(
            Name="Panel15").PaneControl(Name="CONTAINER LABEL").GetChildren()[1]
        order_box.SendKeys(value, waitTime=0)

    def customer(self, value: str) -> None:
        """set customer (NV)"""
        customer_box = self.main_win.PaneControl(Name="").PaneControl(Name="Panel14").PaneControl(
            Name="Panel15").PaneControl(Name="CONTAINER LABEL").GetChildren()[0]
        customer_box.SetFocus()
        customer_box.SendKeys(value, waitTime=0)

    def destination(self, value) -> None:
        """Set Destination, because this thing has auto fill and items with same starting character, I input 1 at front to deactive autofill 
        then complete the destination, goback delete the 1"""
        destination_box = self.main_win.PaneControl(Name="").PaneControl(Name="Panel14").PaneControl(
            Name="Panel15").PaneControl(Name="CONTAINER LABEL").GetChildren()[3]
        destination_box.SetFocus()
        destination_box.SendKeys(f"1{value}"+"{Home}{Delete}", waitTime=0)

    def vehicle_no(self, value: str) -> None:
        """set vehicle number (NV DN)"""
        vehicle_no_box = self.main_win.PaneControl(Name="").PaneControl(Name="Panel14").PaneControl(
            Name="Panel15").PaneControl(Name="CONTAINER LABEL").GetChildren()[6]
        vehicle_no_box.SetFocus()
        vehicle_no_box.SendKeys(value, waitTime=0)

    def prepare_data(self, value: str) -> None:
        """carton id data"""
        prepare_data_box = self.main_win.PaneControl(Name="Panel5").PaneControl(
            Name="Panel6").PaneControl(Name="Panel9").GetChildren()[0].PaneControl(Name="Panel11").EditControl(Name="")
        prepare_data_box.SendKeys(value, waitTime=0)

    def start_button(self) -> None:
        """click the start button to start entering, require all other data except prepare_data has been entered"""
        start_button = self.main_win.PaneControl(
            Name="Panel5").PaneControl(Name="Panel7").ButtonControl(Name="Start")
        start_button.Click(simulateMove=False, waitTime=0)

    def kb_send_button(self) -> None:
        """kb send after entering data"""
        kb_send_button = self.main_win.PaneControl(Name="Panel5").PaneControl(
            Name="Panel6").PaneControl(Name="Panel9").GetChildren()[0].PaneControl(Name="").ButtonControl(Name="KB Send")
        kb_send_button.Click(simulateMove=False, waitTime=0)

    def upload_button(self) -> None:
        """upload after kb sending"""
        upload_button = self.main_win.PaneControl(Name="Panel5").PaneControl(
            Name="Panel6").PaneControl(Name="Panel9").GetChildren()[0].PaneControl(Name="").ButtonControl(Name="Upload")
        upload_button.Click(simulateMove=False, waitTime=0)

    def link_model(self, NV_model: str) -> None:
        self.main_win.SendKeys("{Alt}o")
        link_model_button = self.main_win.MenuControl(
            Name="Option").MenuItemControl(Name="Link Model")
        link_model_button.Click(simulateMove=False)

        model_data = NV_model.split("-")
        FOX_model = f"{model_data[0]}{model_data[1]}-{model_data[2]}{model_data[3]}P"
        time.sleep(1)
        old_model_name_box = uiautomation.WindowControl(Name="Link Model").PaneControl(
            Name="Link Model Info.").ComboBoxControl(Name="").EditControl(Name="")
        new_model_name_box = uiautomation.WindowControl(Name="Link Model").PaneControl(
            Name="Link Model Info.").EditControl(Name="")
        link_button = uiautomation.WindowControl(Name="Link Model").PaneControl(
            Name="Link Model Info.").ButtonControl(Name="^:^  Link")
        search_buton = uiautomation.WindowControl(Name="Link Model").PaneControl(
            Name="Link Model Info.").ButtonControl(Name="Search")
        delete_button = uiautomation.WindowControl(Name="Link Model").PaneControl(
            Name="Link Model Info.").ButtonControl(Name="Delete")
        close_button = uiautomation.WindowControl(
            Name="Link Model").ButtonControl(Name="@ Close")

        old_model_name_box.SendKeys(NV_model)
        new_model_name_box.SendKeys(FOX_model, waitTime=0)
        search_buton.Click(simulateMove=False, waitTime=0)
        delete_button.Click(simulateMove=False, waitTime=0)
        link_button.Click(simulateMove=False)

        link_complete_button = uiautomation.WindowControl(
            Name="Shipping").ButtonControl(Name="OK")
        link_complete_button.Click(simulateMove=False, waitTime=0)
        close_button.Click(simulateMove=False)
