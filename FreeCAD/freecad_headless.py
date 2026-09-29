"""
FreeCAD Headless Module
-----------------------
Runs FreeCAD scripts in a separate headless `freecadcmd` process.
Heavy OpenCASCADE (OCCT) work (helical sweeps, lofts, complex booleans) can segfault
in native C++ code. Inside the GUI process that would crash FreeCAD and lose unsaved work;
here, it executes in an isolated process with exit status, output, and crash signal detection.
"""

import logging
import math
import os
import re
import shlex
import shutil
import signal
import subprocess
import tempfile
from pathlib import Path
from typing import Any, List, Optional, Dict

logger = logging.getLogger("FreeCADHeadless")

FLATPAK_APP = "org.freecad.FreeCAD"
_NOISE = ("%)", "Importing project files", "Postprocessing", "FreeCAD 1.", "(C) 2001", "LGPL")

COMMON_WIN_PATHS = [
    os.environ.get("FREECAD_CMD", ""),
    os.environ.get("FREECAD_PATH", ""),
    os.environ.get("FREECAD_BIN", ""),
    r"C:\Program Files\FreeCAD",
    r"C:\Program Files\FreeCAD 1.1",
    r"C:\Program Files\FreeCAD 1.0",
    r"C:\Program Files\FreeCAD 0.22",
    r"C:\Program Files\FreeCAD 0.21",
    r"C:\Program Files\FreeCAD 0.20",
    r"C:\Program Files (x86)\FreeCAD",
    r"C:\Program Files (x86)\FreeCAD 1.1",
    r"C:\Program Files (x86)\FreeCAD 1.0",
    r"C:\Program Files (x86)\FreeCAD 0.21",
    r"D:\Program Files\FreeCAD",
    r"D:\Program Files\FreeCAD 1.1",
    r"D:\Program Files\FreeCAD 1.0",
    r"D:\Program Files\FreeCAD 0.21",
    r"E:\Program Files\FreeCAD",
    r"E:\Program Files\FreeCAD 1.1",
    r"E:\Program Files\FreeCAD 1.0",
    r"E:\Program Files\FreeCAD 0.21",
    os.path.expanduser(r"~\AppData\Local\Programs\FreeCAD 1.1"),
    os.path.expanduser(r"~\AppData\Local\Programs\FreeCAD 1.0"),
    os.path.expanduser(r"~\AppData\Local\Programs\FreeCAD"),
    os.path.expanduser(r"~\AppData\Local\Programs\FreeCAD 0.21"),
]


def detect_freecadcmd() -> Optional[List[str]]:
    """Locate a headless FreeCAD executable: PATH first, known Windows directories, then Flatpak."""
    # 1. Custom ENV var
    env_cmd = os.environ.get("FREECAD_CMD")
    if env_cmd and os.path.exists(env_cmd):
        return [env_cmd]

    # 2. System PATH
    for name in ("freecadcmd", "FreeCADCmd", "freecadcmd.exe", "FreeCADCmd.exe", "freecad.cmd", "FreeCAD.exe", "freecad.exe"):
        path = shutil.which(name)
        if path:
            return [path]

    # 3. Known Windows install paths
    for base in COMMON_WIN_PATHS:
        if not base or not os.path.exists(base):
            continue
        if os.path.isfile(base) and base.lower().endswith(".exe"):
            return [base]
        
        candidates = [
            os.path.join(base, "bin", "FreeCADCmd.exe"),
            os.path.join(base, "bin", "freecadcmd.exe"),
            os.path.join(base, "bin", "FreeCAD.exe"),
            os.path.join(base, "bin", "freecad.exe"),
            os.path.join(base, "FreeCADCmd.exe"),
            os.path.join(base, "FreeCAD.exe"),
        ]
        for c in candidates:
            if os.path.isfile(c):
                return [c]

    # 4. Flatpak on Linux
    flatpak = shutil.which("flatpak")
    if flatpak:
        try:
            probe = subprocess.run(
                [flatpak, "info", FLATPAK_APP], capture_output=True, timeout=10
            )
            if probe.returncode == 0:
                return [flatpak, "run", "--command=freecadcmd", FLATPAK_APP]
        except (OSError, subprocess.TimeoutExpired):
            pass

    return None


def parse_command(value: Optional[str]) -> Optional[List[str]]:
    return shlex.split(value) if value else None


def _script_dir() -> Path:
    d = Path.home() / ".cache" / "freecad-mcp" / "headless"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _clean(output: Optional[Any]) -> str:
    if output is None:
        return ""
    if isinstance(output, bytes):
        output = output.decode("utf-8", errors="replace")
    lines = [ln for ln in str(output).splitlines() if ln.strip() and not any(n in ln for n in _NOISE)]
    return "\n".join(lines)


def run_headless(code: str, timeout: float = 600.0, command: Optional[List[str]] = None) -> Dict[str, Any]:
    """
    Execute python code headlessly with FreeCADCmd in an isolated process.
    Returns success, returncode, output, crashed, and timed_out flags.
    """
    if not math.isfinite(timeout) or timeout <= 0:
        return {"success": False, "error": "timeout must be a positive finite number"}

    cmd = command or detect_freecadcmd()
    if not cmd:
        return {
            "success": False,
            "error": "FreeCADCmd not found: install FreeCAD (https://www.freecad.org/) or set FREECAD_CMD / FREECAD_PATH",
        }

    script_path = None
    try:
        with tempfile.NamedTemporaryFile("w", suffix=".py", dir=_script_dir(), delete=False, encoding="utf-8") as f:
            f.write(code)
            script_path = f.name

        argv = cmd + ["-c", f"exec(open({script_path!r}, encoding='utf-8').read())"]
        proc = subprocess.run(
            argv,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )

        output = _clean(proc.stdout + "\n" + proc.stderr)
        result: Dict[str, Any] = {
            "success": proc.returncode == 0,
            "returncode": proc.returncode,
            "output": output,
        }

        if proc.returncode < 0:
            sig = -proc.returncode
            try:
                name = signal.Signals(sig).name
            except ValueError:
                name = str(sig)
            result["crashed"] = True
            result["error"] = f"Headless FreeCAD crashed with signal {name} (OCCT native crash; GUI unaffected)"
        elif proc.returncode != 0:
            crash = re.search(r"(?m)^Program received signal (SIG[A-Z0-9]+),", output)
            if crash:
                result["crashed"] = True
                result["error"] = (
                    f"Headless FreeCAD reported a crash with {crash.group(1)} "
                    f"(exit code {proc.returncode}; GUI unaffected)"
                )
            else:
                result["error"] = f"Script failed with exit code {proc.returncode}: {output}"

        return result

    except subprocess.TimeoutExpired as e:
        out = "\n".join(part for part in (_clean(e.stdout), _clean(e.stderr)) if part)
        return {
            "success": False,
            "timed_out": True,
            "error": f"Headless FreeCAD script did not finish within {timeout:g} seconds",
            "output": out,
        }
    except OSError as e:
        return {"success": False, "error": f"Could not start headless FreeCAD: {e}"}
    finally:
        if script_path and os.path.exists(script_path):
            try:
                os.unlink(script_path)
            except OSError:
                pass
