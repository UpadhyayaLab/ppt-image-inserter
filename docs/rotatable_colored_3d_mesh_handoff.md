# Handoff: rotatable (optionally color-preserved) 3D meshes in PowerPoint

Self-contained brief for an agent in another codebase. Everything below was
built and **verified end-to-end** on Windows + PowerPoint (Microsoft 365) +
MATLAB R2024b. Paths are absolute on this machine; copy the code out rather than
importing, since you're in a different repo.

---

## What this does

1. Insert a mesh (`.obj`/`.glb`/`.gltf`/`.fbx`/`.stl`/`.ply`/`.3mf`) into a
   `.pptx` as a **native, interactively rotatable 3D model** (drag-handle orbit).
2. Optionally **preserve a scalar coloring** (e.g. a MATLAB curvature map) that a
   plain `.obj` throws away, so the model looks like the original figure instead
   of a grey blob.

---

## The non-obvious facts (the hard part)

1. **`python-pptx` has NO 3D API.** You must drive PowerPoint via COM. The call
   is `Shapes.Add3DModel(FileName, LinkToFile, SaveWithDocument, Left, Top,
   Width, Height)` (positions in points). Windows + installed PowerPoint only.
2. **PowerPoint converts the input mesh itself.** After `Add3DModel`, the source
   is re-encoded to an embedded `ppt/media/model3dN.glb` **plus** a PNG fallback
   render, wired via relationship type
   `http://schemas.microsoft.com/office/2017/06/relationships/model3d`. The deck
   is self-contained; you don't need a glTF/mesh library just to insert.
3. **Interactive rotation is free**; a default opening angle is set via
   `shape.Model3D.RotationX/Y/Z` (degrees; PPT normalizes to 0–360). Also
   available: `IncrementRotationX/Y/Z`, `CameraPositionX/Y/Z`, `FieldOfView`,
   `ResetModel`, `AutoFit`. **`SetCameraFromPreset` does NOT exist** on this build.
4. **Auto-spin (Turntable) is NOT scriptable.** The `MsoAnimEffect` COM enum has
   no 3D members (it stops at path effects ~149; media effects are 83–85), so
   `TimeLine.MainSequence.AddEffect` cannot add it. Hand-authoring the 3D-anim
   OOXML is corruption-prone. Leave Turntable as a manual step (Animations tab →
   Turntable → Timing → Repeat: Until End of Slide).
5. **A plain `.obj` is geometry only.** MATLAB (and most exporters) color a mesh
   at *render* time from a scalar field — that color is never written to the
   `.obj` (no `mtllib`/`usemtl`/`vt`/vertex-color lines). To preserve color you
   must recover the scalar/RGB from wherever it actually lives (for MATLAB: the
   saved `.fig`) and re-bake it onto the mesh.
6. **PowerPoint's 3D renderer HONORS glTF vertex/face colors** (the `COLOR_0`
   attribute). This was the make-or-break unknown; it's true. Bake **per-face
   flat** colors by *unwelding* the mesh (each triangle gets its own 3 vertices,
   all carrying that face's color) — reproduces MATLAB `FaceColor='flat'` exactly.
   BUT see the color-fidelity section below — vertex color alone is not enough.
7. **MATLAB `.fig` is a MAT-file** holding the whole figure. The colored surface
   is a `patch` with `Vertices`, `Faces`, `FaceVertexCData` (scalar or RGB), and
   the axes `Colormap` + `CLim`. Map scalar→RGB the way MATLAB renders scaled
   CData: `idx = clamp(floor((c-CLim(1))/(CLim(2)-CLim(1))*n)+1, 1, n)`.

---

## PowerPoint 3D color fidelity — the two gotchas that wash colors out

Getting `COLOR_0` colors *into* the model is only half the battle; PowerPoint's
3D renderer will otherwise make them pale. Two independent causes, both verified
by inserting and reading the fallback render:

1. **No material → metallic wash.** A glTF primitive with no `material` gets the
   default (`metallicFactor=1.0`) — a fully metallic surface that reflects the
   scene to near-white. **Always attach a matte material** (`metallicFactor=0`,
   `roughnessFactor=1`, `doubleSided=true`).
2. **Bright studio lights → desaturation.** PowerPoint lights every 3D model
   with a fixed bright rig (ambient + 3 point lights, visible in the `am3d`
   block). A *lit* surface clips toward white, so even correct saturated colors
   render pastel. PowerPoint does **not** honor `KHR_materials_unlit` (it still
   lights the model), and post-hoc editing the `am3d` light rig doesn't help
   because `Slide.Export`/reopen serves the *cached* raster (rendered at insert
   time). The reliable fix is to **pre-compensate the vertex colors**: boost
   saturation (~2.0–2.5×, push away from luma) and darken (~0.8×) before baking,
   so the lit result lands vivid. This trades exact colormap values for a visual
   match — right for a cover/aesthetic figure, note it for a quantitative one.

`recolor_glb_for_powerpoint(in, out, saturate, darken)` in `glb.py` does both
(re-emits matte/unlit + optional boost). `build_colored_glb(..., unlit=True)` is
now the default and always attaches the matte material. Tune `saturate`/`darken`
per palette by inserting one frame and eyeballing the fallback render.

## Verification trick (no need to open PowerPoint)

After `Add3DModel` + `SaveAs`, PowerPoint has already rendered the model to
`ppt/media/image1.png` (the fallback). **Unzip the `.pptx` and open that PNG** —
it's PowerPoint's own render of your model, so it confirms geometry AND color
landed. That's how the color pipeline was validated (a half-red/half-blue cube,
then the real nucleus).

---

## COM gotchas

- The output `.pptx` must be **closed in PowerPoint**, and **no orphan
  `POWERPNT` process** may be running, or the build fails with
  `RPC_E_CALL_REJECTED` / "call was rejected by callee". Kill stray POWERPNT
  first (`Get-Process POWERPNT | Stop-Process -Force`) — but note that also
  closes a user's open PowerPoint, so it's a deliberate step, not automatic.
- `LinkToFile=0` (msoFalse), `SaveWithDocument=-1` (msoTrue) → embed. `SaveAs(path, 24)`
  = `ppSaveAsOpenXMLPresentation`.
- For MATLAB `.fig` extraction, run **one** `matlab -batch` for the whole batch;
  MATLAB startup is ~10–15 s and you don't want to pay it per file.

---

## Files to copy (absolute paths on this machine)

Reusable, low-dependency:

- **`K:\FF\Code\Scripts\PPT_image_inserter\ppt_image_inserter\glb.py`**
  `build_colored_glb(vertices, faces, colors, out_path, per_face=True)` — writes a
  colored `.glb` with `POSITION` + `COLOR_0`. **numpy-only, no trimesh/pygltflib.**
  Self-contained; copy as-is.
- **`K:\FF\Code\Scripts\PPT_image_inserter\ppt_image_inserter\models3d.py`**
  `Model3DSpec`, `Model3DSlideSpec`, `build_model3d_deck_via_com(...)` — builds a
  deck of textboxes + 3D models by emitting a JSON manifest and running a
  PowerShell/COM script. The exact `Add3DModel` + rotation call lives in the
  `_POWERSHELL_TEMPLATE` here. (Imports `TextboxSpec` from `movies.py` — a tiny
  dataclass; copy that too.)
- **`K:\FF\Code\Scripts\PPT_image_inserter\ppt_image_inserter\movies.py`**
  `TextboxSpec` dataclass + `build_movie_deck_via_com` — the COM/PowerShell
  pattern `models3d.py` mirrors, and the reference for autoplay/loop OOXML
  rewriting if you ever embed video.
- **`K:\FF\Code\Scripts\PPT_image_inserter\examples_and_configs\matlab\extract_fig_mesh_color.m`**
  and **`...\extract_fig_mesh_color_batch.m`** — openfig → find the mesh patch →
  reproduce displayed RGB (colormap+CLim) → save `V/F/RGB/perface` to a `-v7`
  `.mat`. The batch driver reads a tab-separated `fig<TAB>outMat` manifest.

Full worked example (experiment-specific glue, but shows the whole flow
`.fig → MATLAB → .mat → build_colored_glb → Add3DModel`):

- **`K:\FF\Code\Scripts\PPT_image_inserter\examples_and_configs\insert_jurkat_nuc_colored_mesh_slides.py`**
- Plain (uncolored) version: **`...\insert_mesh_obj_3d_slides.py`**

---

## The pipeline in one picture

```
MATLAB .fig ──openfig──▶ V, F, FaceVertexCData, Colormap, CLim
                              │  (extract_fig_mesh_color.m)
                              ▼  map scalar→RGB
                         V, F, RGB  (.mat)
                              │  (build_colored_glb: unweld per-face)
                              ▼
                        colored .glb  (glTF COLOR_0)
                              │  (Add3DModel via COM)
                              ▼
                 rotatable colored 3D model in .pptx
```

For an already-colored source (`.glb`/`.ply`/`.fbx` that carries color), skip the
first three steps and just `Add3DModel` it.

---

## Minimal COM call (PowerShell), if you're not in Python

```powershell
$app = New-Object -ComObject PowerPoint.Application
$app.Visible = -1
$pres = $app.Presentations.Add()
$slide = $pres.Slides.Add(1, 12)                 # 12 = ppLayoutBlank
$shape = $slide.Shapes.Add3DModel($glbPath, 0, -1, $leftPt, $topPt, $wPt, $hPt)
$shape.Model3D.RotationX = 20; $shape.Model3D.RotationY = -30
$pres.SaveAs($outPptx, 24)                        # 24 = ppSaveAsOpenXMLPresentation
$pres.Close(); $app.Quit()
```

---

## Dependencies

- Insert path: Windows, PowerPoint installed (COM). Python side uses only stdlib
  + `numpy` (glb writer) + `scipy` (`loadmat`). No trimesh/pygltflib on this box.
- Color extraction: MATLAB (any recent; R2024b here at
  `C:\Program Files\MATLAB\R2024b\bin\matlab.exe`). Needed only because the color
  source is `.fig`; a different color source (`.ply` with vertex colors, a `.mat`
  of per-vertex scalars, etc.) skips MATLAB entirely.
```
