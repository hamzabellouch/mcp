"""
Unit & Integration Verification Tests for FreeCAD MCP Server
------------------------------------------------------------
Tests server initialization, tool schema registration, headless detection,
and FreeCAD engine execution using standard Python unittest.
"""

import sys
import os
import asyncio
import unittest
import tempfile

# Ensure project root is in sys.path
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

import freecad_mcp_server
from freecad_headless import detect_freecadcmd, parse_command, _clean
from freecad_engine import FreeCADEngine


class TestFreeCADMCP(unittest.TestCase):
    """Test suite verifying MCP server schemas, tools, and engine execution."""

    def test_server_instantiation(self):
        """Verify FastMCP server instance exists and has valid name and instructions."""
        mcp = freecad_mcp_server.mcp
        self.assertIsNotNone(mcp)
        self.assertEqual(mcp.name, "FreeCADMCP")
        self.assertIn("FreeCAD", mcp.instructions)

    def test_registered_tool_catalog(self):
        """Verify all 29 tools are registered on the FastMCP instance."""
        expected_tools = [
            # Session & Document Tools
            "create_document",
            "list_documents",
            "reload_document",
            "freecad_open_document",
            "freecad_save_document",
            "freecad_status",
            "get_rpc_status",
            
            # Object Manipulation & View
            "create_object",
            "edit_object",
            "delete_object",
            "get_objects",
            "get_object",
            "get_view",
            "freecad_analyze_geometry",
            
            # High-Level Parametric Tools
            "freecad_create_primitive",
            "freecad_boolean_operation",
            "freecad_apply_fillet_chamfer",
            "freecad_create_sketch_and_extrude",
            "freecad_revolve_shape",
            "insert_part_from_library",
            "get_parts_list",
            
            # Simulation, TechDraw & I/O
            "run_fem_analysis",
            "freecad_create_techdraw",
            "freecad_export_file",
            "freecad_import_file",
            
            # Script Execution
            "execute_code",
            "execute_code_async",
            "get_async_status",
            "execute_code_headless",
        ]
        
        for tool_name in expected_tools:
            self.assertTrue(hasattr(freecad_mcp_server, tool_name), f"Missing tool in module: {tool_name}")
            func = getattr(freecad_mcp_server, tool_name)
            self.assertTrue(callable(func), f"Tool {tool_name} is not callable")

    def test_fastmcp_list_tools_async(self):
        """Test FastMCP tool listing async interface."""
        async def _run():
            mcp = freecad_mcp_server.mcp
            tools = await mcp.list_tools()
            tool_names = [t.name for t in tools]
            
            self.assertIn("freecad_open_document", tool_names)
            self.assertIn("freecad_save_document", tool_names)
            self.assertIn("create_document", tool_names)
            self.assertIn("freecad_create_primitive", tool_names)
            self.assertIn("execute_code", tool_names)
            self.assertGreaterEqual(len(tool_names), 29)

        asyncio.run(_run())

    def test_registered_prompts(self):
        """Verify prompts registration."""
        self.assertTrue(hasattr(freecad_mcp_server, "asset_creation_strategy"))
        strategy = freecad_mcp_server.asset_creation_strategy()
        self.assertIn("Asset Creation Strategy", strategy)

    def test_headless_helpers(self):
        """Test headless command parsing and output cleaning."""
        cmd = parse_command("freecadcmd --version")
        self.assertEqual(cmd, ["freecadcmd", "--version"])
        self.assertIsNone(parse_command(None))
        
        noisy_output = "Postprocessing...\nLGPL FreeCAD\nActual Output Line\n100%)"
        cleaned = _clean(noisy_output)
        self.assertIn("Actual Output Line", cleaned)
        self.assertNotIn("Postprocessing", cleaned)

    def test_freecad_engine_status(self):
        """Test engine status detection."""
        status = FreeCADEngine.get_status()
        self.assertIsInstance(status, dict)
        self.assertIn("rpc_connected", status)
        self.assertIn("http_bridge_connected", status)
        self.assertIn("headless_available", status)

    def test_freecad_status_tool(self):
        """Test freecad_status tool output formatting."""
        status_text = freecad_mcp_server.freecad_status()
        self.assertIn("FreeCAD Status Report", status_text)

    def test_open_save_document_tools(self):
        """Verify freecad_open_document and freecad_save_document behavior."""
        # Test nonexistent file handling
        res_nonexistent = freecad_mcp_server.freecad_open_document("nonexistent_model_file.FCStd")
        self.assertIsInstance(res_nonexistent, str)
        self.assertTrue(
            "❌ Failed to open document" in res_nonexistent or "No FreeCAD runtime available" in res_nonexistent,
            f"Unexpected response: {res_nonexistent}"
        )

        # Test save handling when no document is active
        res_save = freecad_mcp_server.freecad_save_document(document_name="NonexistentDoc")
        self.assertIsInstance(res_save, str)
        self.assertTrue(
            "❌ Failed to save document" in res_save or "No FreeCAD runtime available" in res_save,
            f"Unexpected response: {res_save}"
        )

        # If headless FreeCADCmd is available, test end-to-end create, save, and open
        cmd = detect_freecadcmd()
        if cmd:
            with tempfile.TemporaryDirectory() as tmpdir:
                fcstd_path = os.path.join(tmpdir, "test_box.FCStd").replace("\\", "/")
                # 1. Create a box and save to FCStd
                create_script = f"""
import FreeCAD, Part
doc = FreeCAD.newDocument("BoxDoc")
box = doc.addObject("Part::Box", "MyBox")
box.Length = 20
box.Width = 15
box.Height = 10
doc.recompute()
doc.saveAs("{fcstd_path}")
print(f"SAVED:{{doc.Name}}")
"""
                res = FreeCADEngine.execute_code(create_script)
                self.assertTrue(res.get("success"), f"Headless script failed: {res}")
                self.assertTrue(os.path.exists(fcstd_path), f"File was not created at {fcstd_path}")
                
                # 2. Test opening the saved document
                res_open = freecad_mcp_server.freecad_open_document(fcstd_path)
                self.assertIn("✅ Document opened successfully!", res_open)
                self.assertIn("test_box", res_open)
                self.assertIn("Objects count: 1", res_open)


if __name__ == "__main__":
    unittest.main(verbosity=2)
