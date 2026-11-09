"""
Unit and Integration Tests for Blender MCP Server
-------------------------------------------------
Verifies FastMCP tools, argument parsing, error recovery, and communication logic.
"""

import os
import sys
import unittest
import tempfile
import json
from unittest.mock import patch, MagicMock

# Route bytecode cache to Tests/.cache
sys.pycache_prefix = os.path.abspath(os.path.join(os.path.dirname(__file__), ".cache"))

# Ensure Blender project directory is in sys.path
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "Blender"))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

import blender_mcp_server
from blender_mcp_server import (
    mcp,
    blender_status,
    blender_get_scene_summary,
    blender_execute_script,
    blender_create_primitive,
    blender_assign_material,
    blender_clear_scene,
    blender_render_image,
    blender_import_3d_model,
    main,
)


class TestBlenderMCPOffline(unittest.TestCase):
    """Test tools behavior when Blender HTTP bridge is offline."""

    def test_status_when_offline(self):
        result = blender_status()
        self.assertIn("Disconnected", result)
        self.assertIn("Cannot connect to Blender", result)

    def test_get_scene_summary_when_offline(self):
        result = blender_get_scene_summary()
        self.assertIn("Error fetching scene info", result)

    def test_execute_script_when_offline(self):
        result = blender_execute_script("import bpy")
        self.assertIn("Execution Failed", result)

    def test_clear_scene_when_offline(self):
        result = blender_clear_scene()
        self.assertIn("Error clearing scene", result)

    def test_render_image_when_offline(self):
        result = blender_render_image()
        self.assertIn("Render failed", result)


class TestBlenderMCPMockedOnline(unittest.TestCase):
    """Test tool logic and script synthesis against mocked responses."""

    @patch("blender_mcp_server._send_request")
    def test_blender_status_online(self, mock_send):
        mock_send.return_value = {
            "status": "online",
            "service": "Blender MCP Bridge",
            "blender_version": "4.2.0",
            "version": "1.0.0"
        }
        res = blender_status()
        self.assertIn("Connected to Blender 4.2.0", res)
        self.assertIn("online", res)

    @patch("blender_mcp_server._send_request")
    def test_get_scene_summary(self, mock_send):
        mock_data = {
            "scene_name": "Scene",
            "blender_version": "4.2.0",
            "total_objects": 3,
            "objects": [{"name": "Cube", "type": "MESH"}]
        }
        mock_send.return_value = mock_data
        res = blender_get_scene_summary()
        parsed = json.loads(res)
        self.assertEqual(parsed["scene_name"], "Scene")
        self.assertEqual(len(parsed["objects"]), 1)

    @patch("blender_mcp_server._send_request")
    def test_execute_script_success_output(self, mock_send):
        mock_send.return_value = {
            "success": True,
            "stdout": "Cube created successfully\n",
            "stderr": "",
            "error": None
        }
        res = blender_execute_script("bpy.ops.mesh.primitive_cube_add()")
        self.assertIn("Cube created successfully", res)

    @patch("blender_mcp_server._send_request")
    def test_execute_script_error_output(self, mock_send):
        mock_send.return_value = {
            "success": False,
            "stdout": "",
            "stderr": "",
            "error": "SyntaxError: invalid syntax"
        }
        res = blender_execute_script("invalid syntax !!!")
        self.assertIn("Execution Failed", res)
        self.assertIn("SyntaxError", res)

    @patch("blender_mcp_server._send_request")
    def test_create_primitive_primitives(self, mock_send):
        mock_send.return_value = {"success": True, "stdout": "Created cube named Cube_1", "stderr": "", "error": None}
        
        primitives = [
            ("cube", "primitive_cube_add"),
            ("uv_sphere", "primitive_uv_sphere_add"),
            ("ico_sphere", "primitive_ico_sphere_add"),
            ("cylinder", "primitive_cylinder_add"),
            ("cone", "primitive_cone_add"),
            ("torus", "primitive_torus_add"),
            ("plane", "primitive_plane_add"),
            ("monkey", "primitive_monkey_add"),
            ("light", "light_add"),
            ("camera", "camera_add"),
        ]

        for ptype, expected_func in primitives:
            res = blender_create_primitive(
                primitive_type=ptype,
                name=f"Custom_{ptype}",
                location=[1.0, 2.0, 3.0],
                rotation=[0.0, 0.5, 0.0],
                scale=[2.0, 2.0, 2.0],
                size=3.0
            )
            self.assertTrue(mock_send.called)
            sent_code = mock_send.call_args[1]["payload"]["code"]
            self.assertIn(expected_func, sent_code)
            self.assertIn(f'"Custom_{ptype}"', sent_code)

    def test_create_primitive_unsupported(self):
        res = blender_create_primitive(primitive_type="spaceship_hyperdrive")
        self.assertIn("Unsupported primitive type", res)

    @patch("blender_mcp_server._send_request")
    def test_create_primitive_quote_escaping(self, mock_send):
        mock_send.return_value = {"success": True, "stdout": "OK", "stderr": "", "error": None}
        name_with_quotes = 'Object "Special" with \'quotes\' & \n newlines'
        blender_create_primitive("cube", name=name_with_quotes)
        
        sent_code = mock_send.call_args[1]["payload"]["code"]
        compiled = compile(sent_code, "<test>", "exec")
        self.assertIsNotNone(compiled)

    @patch("blender_mcp_server._send_request")
    def test_assign_material_code_synthesis(self, mock_send):
        mock_send.return_value = {"success": True, "stdout": "Material assigned", "stderr": "", "error": None}
        
        res = blender_assign_material(
            object_name="Suzanne's Head",
            material_name="Gold_Mat\"PBR\"",
            base_color=[1.0, 0.84, 0.0, 1.0],
            metallic=1.0,
            roughness=0.1,
            emission_color=[0.1, 0.1, 0.0, 1.0],
            emission_strength=2.0
        )
        self.assertTrue(mock_send.called)
        sent_code = mock_send.call_args[1]["payload"]["code"]
        compiled = compile(sent_code, "<test_mat>", "exec")
        self.assertIsNotNone(compiled)
        self.assertIn("BSDF_PRINCIPLED", sent_code)
        self.assertIn("[1.0, 0.84, 0.0, 1.0]", sent_code)

    @patch("blender_mcp_server._send_request")
    def test_clear_scene(self, mock_send):
        mock_send.return_value = {"success": True, "deleted_count": 4}
        res = blender_clear_scene(keep_camera=True, keep_lights=False)
        self.assertIn("Removed 4 objects", res)
        mock_send.assert_called_with("/clear", method="POST", payload={
            "keep_camera": True,
            "keep_lights": False
        })

    @patch("blender_mcp_server._send_request")
    def test_render_image(self, mock_send):
        mock_send.return_value = {"success": True, "filepath": "/tmp/test.png"}
        res = blender_render_image(output_path="/tmp/test.png", resolution_x=1920, resolution_y=1080)
        self.assertIn("Render completed successfully", res)
        self.assertIn("/tmp/test.png", res)

    @patch("blender_mcp_server._send_request")
    def test_import_3d_model_local_file(self, mock_send):
        mock_send.return_value = {"success": True, "stdout": "Imported model", "stderr": "", "error": None}
        
        with tempfile.NamedTemporaryFile(suffix=".glb", delete=False) as f:
            temp_model_path = f.name
            
        try:
            res = blender_import_3d_model(
                source_path_or_url=temp_model_path,
                name="TestModel",
                location=[0.0, 1.0, 0.0]
            )
            self.assertTrue(mock_send.called)
            sent_code = mock_send.call_args[1]["payload"]["code"]
            compiled = compile(sent_code, "<test_import>", "exec")
            self.assertIsNotNone(compiled)
            self.assertIn("import_scene.gltf", sent_code)
        finally:
            if os.path.exists(temp_model_path):
                os.remove(temp_model_path)

    def test_import_3d_model_nonexistent_file(self):
        res = blender_import_3d_model(source_path_or_url="/path/to/nonexistent/model.fbx")
        self.assertIn("Model file not found", res)

    @patch("requests.get")
    @patch("blender_mcp_server._send_request")
    def test_import_3d_model_remote_url(self, mock_send, mock_requests_get):
        mock_resp = MagicMock()
        mock_resp.status_code = 200
        mock_resp.iter_content = MagicMock(return_value=[b"dummy glb bytes"])
        mock_requests_get.return_value = mock_resp
        
        mock_send.return_value = {"success": True, "stdout": "Imported from URL", "stderr": "", "error": None}
        
        res = blender_import_3d_model(
            source_path_or_url="https://raw.githubusercontent.com/KhronosGroup/glTF-Sample-Models/master/2.0/Box/glTF-Binary/Box.glb",
            name="DownloadedBox"
        )
        self.assertTrue(mock_requests_get.called)
        self.assertTrue(mock_send.called)


class TestFastMCPIntegration(unittest.TestCase):
    """Test FastMCP server setup, tool registrations, and entrypoint."""

    def test_server_instance(self):
        self.assertEqual(mcp.name, "blender-mcp-server")

    def test_main_function_callable(self):
        self.assertTrue(callable(main))


if __name__ == "__main__":
    unittest.main()
