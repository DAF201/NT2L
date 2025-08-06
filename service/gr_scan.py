import core.SFC
import config.regex_reference
import core.REGEX
import core.TOOL


def get_wo_from_sn(sn: int | str) -> str | None:
    """search for working order number by a sn, make scanning easier"""
    try:
        return core.SFC.global_SFC_API.sn_lookup(sn)["MO_Number"]
    except:
        return None


def get_pn_from_sn(sn: int | str) -> str | None:
    """search for a part number by a sn"""
    try:
        return core.SFC.global_SFC_API.sn_lookup(sn)["Model_Name"]
    except:
        return None


def get_sn_range_from_sn(sn: int | str) -> list[str, str] | None:
    """search for sn range by a sn"""
    query = core.SFC.global_SFC_API.mo_query(get_pn_from_sn(sn))
    wo = get_wo_from_sn(sn)
    for res in query:
        if res["Mo_Number"] == wo:
            return res["SN_Start"], res["SN_End"]
    return None


def get_sn_from_QR() -> list[str] | None:
    """for gi to wo, can just can the QR which will be easier"""
    res = set()
    scanned_qr = set()
    while True:
        core.TOOL.focus_console()
        qr_input = input()
        if qr_input == "quit":
            return None
        if qr_input == "":
            break
        if core.REGEX.re_compare(config.regex_reference.QR_REGEX, qr_input):
            if qr_input not in scanned_qr:
                scanned_qr.add(qr_input)
            else:
                core.TOOL.alert_and_beep("This QR has beed scanned")
                continue
            qr_data = qr_input.split("|")
            for i in range(2, len(qr_data)):
                if qr_data[i] != "":
                    res.add(qr_data[i])
        else:
            core.TOOL.alert_and_beep("Invalid QR")
    return res


def get_carton_id_from_QR() -> list[str] | None:
    """scan a QR to get carton id from it"""
    res = set()
    scanned_qr = set()
    while True:
        core.TOOL.focus_console()
        qr_input = input()
        if qr_input == "quit":
            return None
        if qr_input == "":
            break
        if core.REGEX.re_compare(config.regex_reference.QR_REGEX, qr_input):
            if qr_input not in scanned_qr:
                scanned_qr.add(qr_input)
            else:
                core.TOOL.alert_and_beep("This QR has beed scanned")
                continue
            qr_data = qr_input.split("|")
            res.add(qr_data[1])
        else:
            core.TOOL.alert_and_beep("Invalid QR")
    return res


def scan_one_sn() -> str | None:
    """scan a sn and check if it is valid"""
    while True:
        core.TOOL.focus_console()
        sn_input = input()
        if sn_input == "quit":
            return None
        elif core.REGEX.re_compare(config.regex_reference.SERIAL_REGEX, sn_input):
            return sn_input
        elif sn_input == "":
            return None
        else:
            core.TOOL.alert_and_beep("INVALID SERIAL NUMBER")


def scan_sn() -> list[str]:
    """scan sn, include boundary check and type check"""
    scan_res = set()
    sn_start = 0
    sn_end = 0
    sn_input = scan_one_sn()
    if sn_input == None:
        return []
    sn_start, sn_end = get_sn_range_from_sn(sn_input)
    scan_res.add(sn_input)
    sn_start = int(sn_start)
    sn_end = int(sn_end)
    while True:
        core.TOOL.focus_console()
        sn_input = input()
        if sn_input == "quit":
            return
        if core.REGEX.re_compare(config.regex_reference.SERIAL_REGEX, sn_input):
            sn_input = int(sn_input)
            if sn_input < sn_start or sn_input > sn_end:
                core.TOOL.alert_and_beep("SERIAL NUMBER NOT IN RANGE")
                continue
            else:
                if str(sn_input) not in scan_res:
                    scan_res.add(str(sn_input))
                else:
                    core.TOOL.alert_and_beep(
                        f"Serial Number: {sn_input} has been scanned")
        # user may want to finish or check number of units scanned
        elif sn_input == "":
            # show how many units has been scanned
            core.TOOL.message(
                f"{len(scan_res)} scanned, it this the correct number? (Press 'Enter' to finish)")
            user_comfirmation = input()
            # if enter quit, just quit and do not do anything
            if user_comfirmation == "quit":
                return
            # if enter, finish scanning and return result
            elif user_comfirmation == "":
                core.TOOL.message("scanning complete")
                break
            # if scanned a SN, check if this is valid or not, if yes add to the scanned list, else continue
            elif core.REGEX.re_compare(config.regex_reference.SERIAL_REGEX, user_comfirmation):
                user_comfirmation = int(user_comfirmation)
                if user_comfirmation < sn_start or user_comfirmation > sn_end:
                    core.TOOL.message("SN not in range, scanning continue...")
                else:
                    if str(user_comfirmation) not in scan_res:
                        scan_res.add(str(user_comfirmation))
                        core.TOOL.message("SN added, scanning continue...")
            # something else not a SN, continue
            else:
                core.TOOL.message("scanning continue...")
        else:
            core.TOOL.alert_and_beep("INVALID SERIAL NUMBER")
    return list(scan_res)
