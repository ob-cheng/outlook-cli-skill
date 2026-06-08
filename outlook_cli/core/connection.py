"""Outlook COM connection management."""

import win32com.client as win32


def connect_to_outlook():
    """Establish connection to Outlook via COM.

    Returns:
        tuple: (outlook_app, mapi_namespace)

    Raises:
        Exception: If Outlook is not installed or not running
    """
    try:
        outlook = win32.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace("MAPI")
        return outlook, namespace
    except Exception as e:
        raise RuntimeError(
            f"Failed to connect to Outlook: {e}\n"
            "Make sure Outlook is installed and configured."
        )
