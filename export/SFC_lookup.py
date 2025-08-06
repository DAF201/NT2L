import core.SFC
import service.gr_scan
import core.REGEX
import config.regex_reference


def SFC_product_tracking_by_sn(sn: int | str) -> dict | None:
    """return the tracking infomation about a serial number"""
    if sn == "quit" or core.REGEX.re_compare(config.regex_reference.SERIAL_REGEX, str(sn)) is False:
        return None
    return core.SFC.global_SFC_API.sn_lookup(sn)


def SFC_OQC_cartoon_id_list_by_wo(work_order: str) -> list[str] | None:
    """return the cartoon id in OQC by work order"""
    if work_order == "" or work_order == "quit":
        return None
    return core.SFC.global_SFC_API.oqc_tracking(work_order)


def SFC_PACKING_serial_number_list_by_wo(work_order: str) -> list[str] | None:
    """return the serial numbers in Packing for GI to WO"""
    if work_order == "" or work_order == "quit":
        return None
    return core.SFC.global_SFC_API.packing_tracking(work_order)


def SFC_wo_query(work_order: str, department="") -> list[dict] | None:
    """return the information about a wo in one/all department(s)"""
    if work_order == "" or work_order == "quit":
        return None
    return core.SFC.global_SFC_API.wo_query(work_order, department)


def SFC_mo_query(part_number: str, work_order="") -> list[dict] | None:
    """return the information of all/this working order based on part number"""
    if part_number == "" or part_number == "quit":
        return None
    if work_order == "":
        return core.SFC.global_SFC_API.mo_query(part_number)
    else:
        for res in core.SFC.global_SFC_API.mo_query(part_number):
            if res["Mo_Number"] == work_order:
                return [res]


def SFC_wo_lookup_by_sn(sn: int | str) -> str | None:
    """return working order number of a sn"""
    if sn == "quit" or core.REGEX.re_compare(config.regex_reference.SERIAL_REGEX, str(sn)) is False:
        return None
    return service.gr_scan.get_wo_from_sn(sn)


def SFC_range_lookup_by_sn(sn: int | str) -> list[str, str]:
    """return the start and end of a sn range this sn belongs to"""
    if sn == "quit" or core.REGEX.re_compare(config.regex_reference.SERIAL_REGEX, str(sn)) is False:
        return None
    return service.gr_scan.get_sn_range_from_sn(sn)


def SFC_pn_lookup_by_sn(sn: int | str) -> str:
    """return the part number of this sn"""
    if sn == "quit" or core.REGEX.re_compare(config.regex_reference.SERIAL_REGEX, str(sn)) is False:
        return None
    return service.gr_scan.get_pn_from_sn(sn)


def SFC_OQC_carton_id_list_lookup_by_sn(sn: int | str) -> list[str]:
    """return the carton ids in OQC in this wo with this sn"""
    if sn == "quit" or core.REGEX.re_compare(config.regex_reference.SERIAL_REGEX, str(sn)) is False:
        return None
    return SFC_OQC_cartoon_id_list_by_wo(SFC_wo_lookup_by_sn(sn))


def SFC_PACKING_sn_list_lookup_by_sn(sn: int | str):
    """return the sns in Packing in this wo with this sn"""
    if sn == "quit" or core.REGEX.re_compare(config.regex_reference.SERIAL_REGEX, str(sn)) is False:
        return None
    return SFC_PACKING_serial_number_list_by_wo(SFC_wo_lookup_by_sn(sn))
