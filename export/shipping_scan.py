import service.SAP_shipping_scan
import functions.target_table


def shipping_scan_NVDN(NVDN: str, carton_ids: list[str]) -> None:
    """shipping scan for the NVDN"""
    target_table = functions.target_table.target_table()
    dn_info = target_table.search_dn_for_info(NVDN)
    service.SAP_shipping_scan.SAP_shipping_scan_NVDN(
        NVDN, dn_info["FX DN"], dn_info["PN#"], service.SAP_shipping_scan.destination_convert(dn_info["Ship to"].lower()), "\n".join(carton_ids))


def shipping_scan_FXSJ(row_index: int, carton_ids: list[str]) -> None:
    """shipping scan for the FXSJ vender pool"""
    target_table = functions.target_table.target_table()
    dn_info = target_table.search_row_for_info(row_index)
    service.SAP_shipping_scan.SAP_shipping_scan_FXSJ(
        dn_info["FX DN"], dn_info["PN#"], "USA", "\n".join(carton_ids))
