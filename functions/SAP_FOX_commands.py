import core.SAP_FOX


def change_page(session: core.SAP_FOX.SAPSession, command: str) -> None | str:
    """change the current working page"""
    try:
        command_field = session[core.SAP_FOX.SAPSession.id_command_field]
        command_field.text = command
        session[core.SAP_FOX.SAPSession.id_main_pannel].enter()
        return command
    except:
        return None


def save_page(session: core.SAP_FOX.SAPSession) -> None | bool:
    """press save button"""
    try:
        save_button = session[core.SAP_FOX.SAPSession.id_save_button]
        save_button.press()
        return True
    except:
        return None


def last_process(session: core.SAP_FOX.SAPSession) -> None | bool:
    """go back to last SAP process"""
    try:
        back_button = session[core.SAP_FOX.SAPSession.id_backward_button]
        back_button.press()
        return True
    except:
        return None


def get_widget_value(session: core.SAP_FOX.SAPSession, id: str) -> str | None:
    """get the text value of a widget"""
    try:
        return session[id].text
    except:
        return None


def set_widget_value(session: core.SAP_FOX.SAPSession, id: str, value: str) -> None | bool:
    try:
        session[id].text = value
        return True
    except:
        return None


def press_widget(session: core.SAP_FOX.SAPSession, id: str) -> None | bool:
    try:
        session[id].press()
        return True
    except:
        return None


def press_enter(session: core.SAP_FOX.SAPSession) -> None | bool:
    try:
        session[core.SAP_FOX.SAPSession.id_main_pannel].enter()
        return True
    except:
        return None


def quit_SAP(session: core.SAP_FOX.SAPSession) -> None | bool:
    """close existing SAP session"""
    try:
        command_field = session[core.SAP_FOX.SAPSession.id_command_field]
        command_field.text = "/nex"
        session[core.SAP_FOX.SAPSession.id_main_pannel].enter()
        return True
    except:
        return None


def dump_children(obj: core.SAP_FOX.SAPWidget, indent=0) -> None:
    """to print out all the children widgets on current page recursively"""
    try:
        count = obj.widget.Children.Count
        for i in range(count):
            child = obj.widget.Children(i)
            print("  " * indent + f"{child.Id} - {child.Type}")
            dump_children(core.SAP_FOX.SAPWidget(child), indent + 1)
    except:
        pass
