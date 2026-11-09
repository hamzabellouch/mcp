import sys as _sys
import os as _os
try:
    _addon_dir = _os.path.dirname(_os.path.abspath(__file__))
except NameError:
    import inspect as _inspect
    _addon_dir = _os.path.dirname(_os.path.abspath(_inspect.getfile(_inspect.currentframe())))
if _addon_dir not in _sys.path:
    _sys.path.insert(0, _addon_dir)


class FreeCADMCPAddonWorkbench(Workbench):
    MenuText = "FreeCAD MCP"
    ToolTip = "Workbench for MCP Communication"

    def Initialize(self):
        from rpc_server import rpc_server

        commands = [
            "Start_RPC_Server",
            "Toggle_Auto_Start",
            "Toggle_Remote_Connections",
            "Configure_Allowed_IPs",
            "Set_Auth_Token",
        ]
        self.appendToolbar("FreeCAD MCP", commands)
        self.appendMenu("FreeCAD MCP", commands)

    def Activated(self):
        pass

    def Deactivated(self):
        pass

    def ContextMenu(self, recipient):
        pass

    def GetClassName(self):
        return "Gui::PythonWorkbench"


Gui.addWorkbench(FreeCADMCPAddonWorkbench())


def _auto_start_mcp():
    try:
        from rpc_server import rpc_server

        settings = rpc_server.load_settings()
        if not settings.get("auto_start_rpc", True):
            return

        msg = rpc_server.start_rpc_server()
        FreeCAD.Console.PrintMessage(f"[MCP] Auto-start: {msg}\n")
        try:
            from rpc_server.gui import _update_server_action_ui
            _update_server_action_ui(True)
        except Exception:
            pass
    except Exception as e:
        FreeCAD.Console.PrintWarning(f"[MCP] Auto-start failed: {e}\n")


def _setup_global_ui():
    try:
        from rpc_server import gui as rpc_gui
        rpc_gui.setup_global_menu()
    except Exception as e:
        import FreeCAD
        FreeCAD.Console.PrintWarning(f"[MCP] Failed to install global menu: {e}\n")


try:
    from PySide import QtCore
except ImportError:
    try:
        from PySide2 import QtCore
    except ImportError:
        from PySide6 import QtCore

QtCore.QTimer.singleShot(0, _auto_start_mcp)
QtCore.QTimer.singleShot(600, _setup_global_ui)
