"""
FreeCAD MCP RPC Server Package (Consolidated Modular Architecture)
=================================================================
This package is structured into 3 cohesive modules:
  - server.py: XML-RPC server, IP filter, security & settings
  - geometry.py: 3D Object creation, property mapping, validation & parts
  - gui.py: GUI-thread dispatching, screenshot capture & workbench commands
"""

from . import server
from . import geometry
from . import gui

# Backward-compatibility alias so existing imports like `from rpc_server import rpc_server`
# and `rpc_server.load_settings()` work seamlessly.
rpc_server = server
