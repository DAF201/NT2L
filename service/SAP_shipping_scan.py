import core.SAP_SHIPPING
import config.config


def destination_convert(dest: str) -> str:

    dest_map = {"sc": "USA", "hk": "HONGKONG",
                "india": "INDIA", "dallas": "USA", "tw": "Taiwan", "or": "USA",
                "china": "CHINA", "israel": "Israel","cn":"CHINA"}
    if dest in dest_map.keys():
        return dest_map[dest]
    else:
        return ""


def SAP_shipping_scan_NVDN(NVDN, FXDN, FXPN, dest, carton_ids):
    sap_shipping = core.SAP_SHIPPING.SAP_Shipping()
    sap_shipping.operator_id(
        config.config.global_config["SAP Shipping"]["operator_id"])
    sap_shipping.link_model(FXPN)
    sap_shipping.order(str(int(FXDN)))
    sap_shipping.customer("NV")
    sap_shipping.destination(dest)
    sap_shipping.vehicle_no(str(int(NVDN)))
    sap_shipping.start_button()
    sap_shipping.prepare_data(carton_ids)
    sap_shipping.kb_send_button()


def SAP_shipping_scan_FXSJ(FXDN, FXPN, dest, carton_ids):
    sap_shipping = core.SAP_SHIPPING.SAP_Shipping()
    sap_shipping.operator_id(
        config.config.global_config["SAP Shipping"]["operator_id"])
    sap_shipping.link_model(FXPN)
    sap_shipping.order(str(int(FXDN)))
    sap_shipping.customer("NV")
    sap_shipping.destination(dest)
    sap_shipping.vehicle_no(str(int(FXDN)))
    sap_shipping.start_button()
    sap_shipping.prepare_data(carton_ids)
    sap_shipping.kb_send_button()
