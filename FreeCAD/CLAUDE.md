# FreeCAD MCP Guidelines & Operational Rules

## Foundational Directive: Interactive Step-by-Step Visual CAD Protocol

When working with FreeCAD and the user:

### 1. In-Memory Execution (No .py files on disk)
- Execute all CAD scripts, commands, and geometry operations directly in-memory via the live RPC bridge.
- **Never** write, create, or save temporary or scratch `.py` files on disk unless the user explicitly requests a script file.

### 2. Fast Mode & No Screenshots by Default
- Do **not** capture or return viewport screenshots by default.
- Run in high-speed text/geometry verification mode.
- Capture a screenshot and present it **only** when the user explicitly commands it (e.g. "take a picture", "show screenshot", "أرني لقطة").

### 3. Progressive Step-by-Step Construction (Live Animation)
- Never jump straight to the final model or apply changes in a sudden, single unobservable jump.
- Break down any geometric modeling or feature addition into observable visual stages:
  1. **Stage 1 (2D Wireframe/Contours)**: Draw lines, arcs, or points progressively and update the GUI so the user sees the 2D layout in space.
  2. **Stage 2 (Planar Face)**: Generate the 2D face from closed boundaries.
  3. **Stage 3 (3D Solid Extrusion/Loft/Pocketing)**: Extrude or cut in depth stages.
  4. **Stage 4 (Finishing)**: Apply boolean cuts/unions, fillets, chamfers, and realistic material colors/transparencies.

### 4. Live Animation Pacing
- In multi-step Python execution scripts via `execute_code`, include `FreeCADGui.updateGui()` and brief pauses (`time.sleep(0.04)` to `0.15`) between edges or stages so the construction looks like a live CAD animation in FreeCAD.

### 5. Live Thinking & Visible Iteration (Trial-and-Error on Screen)
- Never resolve all design ideas, color palettes, and dimensions in hidden internal thinking and then dump the final answer all at once.
- Externalize your thought process onto the active 3D canvas:
  - If testing colors or materials, apply the candidate color live first, let the user see it, then refine or replace it live.
  - If adjusting sizes or positions, place the draft geometry on the canvas first, then tune and fine-tune it in observable steps.
  - Make the design process feel like an interactive engineer actively experimenting and perfecting the CAD model in their workshop.

### 6. Dynamic Motion (Rotors, Gears, Mechanisms)
- When animating mechanisms, use `QtCore.QTimer` in the GUI thread to run smooth 30-40 FPS animations with safe cleanup guards.

### 7. User Override
- Always prioritize and respect explicit user instructions. If the user asks for a standalone `.py` file, generate it; if the user asks for a screenshot, capture it. Otherwise, strictly adhere to the default in-memory, no-screenshot, step-by-step, visible-iteration rules.
