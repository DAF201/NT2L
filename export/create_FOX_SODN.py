import service.SAP_FOX_SODN
import core.SAP_FOX
import core.TOOL
import pyperclip


def create_FOX_SODN(nvdn: str, shipping_code: str):
    SAP_application = core.SAP_FOX.SAP()
    SAP_session = SAP_application.session
    so = service.SAP_FOX_SODN.create_FOX_SO_for_NV_DN(
        SAP_session, nvdn, shipping_code)
    dn = service.SAP_FOX_SODN.create_FOX_DN_for_FOX_SO(SAP_session, so)
    core.TOOL.message("SO DN has been copied to clipboard")
    pyperclip.copy(f"{so}\t{dn}")
    return so, dn


def create_FXSJ_SODN(row_index: int):
    SAP_application = core.SAP_FOX.SAP()
    SAP_session = SAP_application.session
    so = service.SAP_FOX_SODN.create_FOX_SO_for_FXSJ(SAP_session, row_index)
    dn = service.SAP_FOX_SODN.create_FOX_DN_for_FOX_SO(SAP_session, so)
    core.TOOL.message("SO DN has been copied to clipboard")
    pyperclip.copy(f"{so}\t{dn}")
    return so, dn
