import service.gr_scan


def scan_sn_from_qr() -> list[str] | None:
    """for GI to WO scanning, will print to screen"""
    return list(service.gr_scan.get_sn_from_QR())


def scan_carton_id_from_qr() -> list[str] | None:
    """for shipping scan, will print to screen"""
    return list(service.gr_scan.get_carton_id_from_QR())
