import core.SAP_FOX
import core.TOOL
import functions.SAP_FOX_commands
import functions.target_table


def create_FOX_SO_for_NV_DN(session: core.SAP_FOX.SAPSession, dn: str, shipping_code="SJ03") -> str:
    """create the FOX SO for the NV DN"""

    # change page to /nva01
    functions.SAP_FOX_commands.change_page(
        session, core.SAP_FOX.SAPSession.id_so_page_command)

    # fill those fixed things
    session[core.SAP_FOX.SAPSession.id_so_sales_document_type].text = "ZHMQ"
    session[core.SAP_FOX.SAPSession.id_so_sales_organization].text = "NV05"
    session[core.SAP_FOX.SAPSession.id_so_distribution_channel].text = "E0"
    session[core.SAP_FOX.SAPSession.id_so_division].text = "MB"
    functions.SAP_FOX_commands.press_enter(session)

    # search for data
    target_table = functions.target_table.target_table()
    dn_data = target_table.search_dn_for_info(dn)

    # if data found
    if dn_data != {}:
        # enter data
        # fixed
        session[core.SAP_FOX.SAPSession.id_so_sold_to_party].text = "BNV021"
        session[core.SAP_FOX.SAPSession.id_so_ship_to_party].text = shipping_code
        session[core.SAP_FOX.SAPSession.id_so_cust_refference].text = dn
        pn_chunks = dn_data["PN#"].replace(
            " ", "").split("-")
        session[core.SAP_FOX.SAPSession.id_so_material_number].text = f"{''.join((pn_chunks[0], pn_chunks[1]))}-{''.join((pn_chunks[2], pn_chunks[3]))}P"
        session[core.SAP_FOX.SAPSession.id_so_sale_units].text = dn_data["DN Qty"]
        session[core.SAP_FOX.SAPSession.id_so_material_number_used_by_customer].text = dn_data["PN#"].replace(
            " ", "")
        session[core.SAP_FOX.SAPSession.id_so_condition_amount_or_percentage].text = "10"
        functions.SAP_FOX_commands.save_page(session)
        while session.has_pop_out():
            session[core.SAP_FOX.SAPSession.id_pop_out_close].press()
        functions.SAP_FOX_commands.change_page(session, "/nvl01n")
        return session[core.SAP_FOX.SAPSession.id_dn_sale_document].text
    else:
        core.TOOL.alert("DN not found in target table")


def create_FOX_DN_for_FOX_SO(session: core.SAP_FOX.SAPSession, so: str) -> str:
    """create DN for a SO"""
    functions.SAP_FOX_commands.change_page(session, "/nvl01n")
    session[core.SAP_FOX.SAPSession.id_dn_shipping_receiving_point].text = "UBLH"
    session[core.SAP_FOX.SAPSession.id_dn_sale_document].text = so
    functions.SAP_FOX_commands.press_enter(session)
    session[core.SAP_FOX.SAPSession.id_dn_storage_location].text = "6L62"
    functions.SAP_FOX_commands.save_page(session)
    functions.SAP_FOX_commands.change_page(session, "/nvl02n")
    return session[core.SAP_FOX.SAPSession.id_DN_delivery].text


def create_FOX_SO_for_FXSJ(session: core.SAP_FOX.SAPSession, row_index: int):
    """create FOX SO for a FXSJ vendor pool"""
    functions.SAP_FOX_commands.change_page(
        session, core.SAP_FOX.SAPSession.id_so_page_command)

    session[core.SAP_FOX.SAPSession.id_so_sales_document_type].text = "ZHMQ"
    session[core.SAP_FOX.SAPSession.id_so_sales_organization].text = "NV05"
    session[core.SAP_FOX.SAPSession.id_so_distribution_channel].text = "E0"
    session[core.SAP_FOX.SAPSession.id_so_division].text = "MB"
    functions.SAP_FOX_commands.press_enter(session)

    target_table = functions.target_table.target_table()
    dn_data = target_table.search_row_for_info(row_index)

    if dn_data != {}:
        session[core.SAP_FOX.SAPSession.id_so_sold_to_party].text = "BNV021"
        session[core.SAP_FOX.SAPSession.id_so_ship_to_party].text = "SJ03"
        session[core.SAP_FOX.SAPSession.id_so_cust_refference].text = dn_data["PB#"]
        pn_chunks = dn_data["PN#"].replace(" ", "").split("-")
        session[core.SAP_FOX.SAPSession.id_so_material_number].text = f"{''.join((pn_chunks[0], pn_chunks[1]))}-{''.join((pn_chunks[2], pn_chunks[3]))}P"
        session[core.SAP_FOX.SAPSession.id_so_sale_units].text = dn_data["DN Qty"]
        session[core.SAP_FOX.SAPSession.id_so_material_number_used_by_customer].text = dn_data["PN#"].replace(" ", "")
        session[core.SAP_FOX.SAPSession.id_so_condition_amount_or_percentage].text = "10"
        functions.SAP_FOX_commands.save_page(session)
        while session.has_pop_out():
            session[core.SAP_FOX.SAPSession.id_pop_out_close].press()
        functions.SAP_FOX_commands.change_page(session, "/nvl01n")
        return session[core.SAP_FOX.SAPSession.id_dn_sale_document].text
    else:
        core.TOOL.alert("Row does not exist")


# def create_FOX_DN_for_FXSJ(session: core.SAP_FOX.SAPSession, so: str):
#     """create FOX DN for a FXSJ vendor pool"""
#     functions.SAP_FOX_commands.change_page(session, "/nvl01n")
#     session[core.SAP_FOX.SAPSession.id_dn_shipping_receiving_point].text = "UBLH"
#     session[core.SAP_FOX.SAPSession.id_dn_sale_document].text = so
#     functions.SAP_FOX_commands.press_enter(session)
#     session[core.SAP_FOX.SAPSession.id_dn_storage_location].text = "6L62"
#     functions.SAP_FOX_commands.save_page(session)
#     functions.SAP_FOX_commands.change_page(session, "/nvl02n")
#     return session[core.SAP_FOX.SAPSession.id_DN_delivery].text
