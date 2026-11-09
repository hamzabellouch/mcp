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
from freecad_mcp_server import (
    detect_freecadcmd,
    parse_command,
    _clean,
    FreeCADEngine,
    FreeCADBridgeHTTPHandler,
)


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
            "freecad_get_active_project_info",
            "freecad_apply_color_matrix",
            "freecad_animate_mechanism",
            "freecad_export_production_bundle",
            
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
        res = freecad_mcp_server.freecad_status()
        self.assertIsInstance(res, list)
        self.assertIsInstance(res[0], freecad_mcp_server.TextContent)
        self.assertIn("FreeCAD Status Report", res[0].text)

    def test_menu_commands_checkable(self):
        """Verify Addon menu toggle commands have Checkable=True."""
        commands_file = os.path.join(PROJECT_ROOT, "addon", "FreeCADMCP", "rpc_server", "commands.py")
        with open(commands_file, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn('"Checkable": True', content)
        self.assertIn('"Checked": bool(settings.get("remote_enabled"', content)
        self.assertIn('"Checked": bool(settings.get("auto_start_rpc"', content)

    def test_cad_modeling_and_export_headless(self):
        """End-to-end CAD test: Primitive + Sketch + Pad + Analyze + Export STEP & STL."""
        cmd = detect_freecadcmd()
        if not cmd:
            self.skipTest("Headless FreeCADCmd not available for integration CAD test.")

        with tempfile.TemporaryDirectory() as tmpdir:
            step_file = os.path.join(tmpdir, "model_export.step").replace("\\", "/")
            stl_file = os.path.join(tmpdir, "model_export.stl").replace("\\", "/")
            cad_script = f"""
import FreeCAD, Part, Sketcher, Mesh, json
doc = FreeCAD.newDocument("TestCADDoc")

# 1. Create Primitive Box
box = doc.addObject("Part::Box", "BaseBlock")
box.Length = 40.0
box.Width = 30.0
box.Height = 20.0
doc.recompute()

# 2. Create PartDesign Body and Pad
body = doc.addObject("PartDesign::Body", "BossBody")
sketch = body.newObject("Sketcher::SketchObject", "BossSketch")
if hasattr(body, "Origin") and body.Origin:
    for f in getattr(body.Origin, "OriginFeatures", []):
        if "XY_Plane" in f.Name or "XY_Plane" in f.Label:
            sketch.AttachmentSupport = [(f, "")]
            sketch.MapMode = "FlatFace"
            break

p1 = FreeCAD.Vector(-10, -5, 0)
p2 = FreeCAD.Vector(10, -5, 0)
p3 = FreeCAD.Vector(10, 5, 0)
p4 = FreeCAD.Vector(-10, 5, 0)
i0 = sketch.addGeometry(Part.LineSegment(p1, p2), False)
i1 = sketch.addGeometry(Part.LineSegment(p2, p3), False)
i2 = sketch.addGeometry(Part.LineSegment(p3, p4), False)
i3 = sketch.addGeometry(Part.LineSegment(p4, p1), False)
sketch.addConstraint(Sketcher.Constraint("Coincident", i0, 2, i1, 1))
sketch.addConstraint(Sketcher.Constraint("Coincident", i1, 2, i2, 1))
sketch.addConstraint(Sketcher.Constraint("Coincident", i2, 2, i3, 1))
sketch.addConstraint(Sketcher.Constraint("Coincident", i3, 2, i0, 1))
doc.recompute()

pad = body.newObject("PartDesign::Pad", "Pad_Boss")
pad.Profile = sketch
pad.Length = 15.0
doc.recompute()

# 3. Analyze Geometry
vol = round(box.Shape.Volume, 4)
print(f"VOL:{{vol}}")

# 4. Export STEP
import Import
export_list = [o for o in doc.Objects if hasattr(o, 'Shape') and o.Shape and not o.Shape.isNull()]
Import.export(export_list, "{step_file}")

# 5. Export STL
mesh_objs = []
try:
    for o in export_list:
        m = doc.addObject("Mesh::Feature", f"__temp_mesh_{{o.Name}}")
        m.Mesh = Mesh.Mesh(o.Shape.tessellate(0.5))
        mesh_objs.append(m)
    Mesh.export(mesh_objs, "{stl_file}")
finally:
    for m in mesh_objs:
        try:
            doc.removeObject(m.Name)
        except Exception:
            pass

print("CAD_PIPELINE_COMPLETE")
"""
            res = FreeCADEngine.execute_code(cad_script)
            self.assertTrue(res.get("success"), f"Headless execution failed: {res.get('error')}")
            self.assertIn("CAD_PIPELINE_COMPLETE", res.get("stdout", ""))
            self.assertIn("VOL:24000", res.get("stdout", ""))
            self.assertTrue(os.path.exists(step_file) and os.path.getsize(step_file) > 0)
            self.assertTrue(os.path.exists(stl_file) and os.path.getsize(stl_file) > 0)

    def test_open_save_document_tools(self):
        """Verify freecad_open_document and freecad_save_document behavior."""
        # Test nonexistent file handling
        res_nonexistent = freecad_mcp_server.freecad_open_document("nonexistent_model_file.FCStd")
        self.assertIsInstance(res_nonexistent, list)
        text_nonexistent = res_nonexistent[0].text
        self.assertTrue(
            "❌ Failed to open document" in text_nonexistent or "No FreeCAD runtime available" in text_nonexistent,
            f"Unexpected response: {text_nonexistent}"
        )

        # Test save handling when no document is active
        res_save = freecad_mcp_server.freecad_save_document(document_name="NonexistentDoc")
        self.assertIsInstance(res_save, list)
        text_save = res_save[0].text
        self.assertTrue(
            "❌ Failed to save document" in text_save or "No FreeCAD runtime available" in text_save,
            f"Unexpected response: {text_save}"
        )

    def test_http_bridge_security_guard(self):
        """Verify HTTP bridge rejects requests from web browser Origin headers."""
        class FakeHandler:
            def __init__(self, headers):
                self.headers = headers
        
        # Request with Origin header should be rejected
        handler_with_origin = FakeHandler({"Origin": "https://malicious-site.com", "Host": "127.0.0.1:9877"})
        rejection = FreeCADBridgeHTTPHandler._verify_request(handler_with_origin)
        self.assertIsNotNone(rejection)
        self.assertEqual(rejection[0], 403)
        
        # Request with DNS rebinding host should be rejected
        handler_with_bad_host = FakeHandler({"Host": "evil-domain.com:9877"})
        rejection_host = FreeCADBridgeHTTPHandler._verify_request(handler_with_bad_host)
        self.assertIsNotNone(rejection_host)
        self.assertEqual(rejection_host[0], 403)


if __name__ == "__main__":
    unittest.main(verbosity=2)
