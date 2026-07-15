"""Minimal, dependency-free glTF 2.0 (.glb) writer for colored meshes.

python-pptx has no 3D API and PowerPoint's ``Add3DModel`` converts whatever it
is given to an internal glTF binary. A mesh coloured by a scalar field (e.g. a
MATLAB ``FaceVertexCData`` curvature map) loses its colour when exported as a
plain ``.obj`` (which stores geometry only). This module bakes per-face or
per-vertex colours into a ``.glb`` via the glTF ``COLOR_0`` vertex attribute,
which PowerPoint's 3D renderer honours -- verified by inserting a colored
``.glb`` and reading back the fallback raster PowerPoint renders from it.

Only :mod:`numpy` is required. The writer emits a single-primitive triangle
mesh with ``POSITION`` + ``COLOR_0`` (+ indices) in one binary buffer.
"""
from __future__ import annotations

import json
import struct

import numpy as np

__all__ = ["build_colored_glb", "read_glb_mesh", "recolor_glb_for_powerpoint"]

_LUMA = np.array([0.299, 0.587, 0.114])

# glTF constants
_FLOAT = 5126
_UINT = 5125
_ARRAY_BUFFER = 34962
_ELEMENT_ARRAY_BUFFER = 34963
_TRIANGLES = 4


def _pad4(b: bytes) -> bytes:
    """Right-pad a byte string to a 4-byte boundary with zeros."""
    return b + b"\x00" * ((-len(b)) % 4)


def _material(unlit: bool) -> dict:
    """A matte, double-sided material; optionally KHR_materials_unlit."""
    mat = {
        "pbrMetallicRoughness": {
            "baseColorFactor": [1.0, 1.0, 1.0, 1.0],
            "metallicFactor": 0.0,
            "roughnessFactor": 1.0,
        },
        "doubleSided": True,
    }
    if unlit:
        mat["extensions"] = {"KHR_materials_unlit": {}}
    return mat


def build_colored_glb(vertices, faces, colors, out_path, per_face=True, unlit=True):
    """Write a colored triangle mesh to a binary glTF (``.glb``) file.

    Args:
        vertices: (N,3) array of vertex positions.
        faces: (M,3) array of 0-based triangle vertex indices.
        colors: RGB(A) colours in [0,1]. When ``per_face`` is True this is
            (M,3) or (M,4), one colour per triangle; the mesh is unwelded so
            each triangle carries its own flat colour (matching MATLAB's
            ``FaceColor='flat'``). When False it is (N,3) or (N,4), one colour
            per vertex, interpolated across triangles (``FaceColor='interp'``).
        out_path: destination ``.glb`` path.
        per_face: whether ``colors`` is indexed per-face (default) or
            per-vertex.
        unlit: when True (default), tag the material ``KHR_materials_unlit`` so
            the vertex colours are shown exactly as authored, independent of
            scene lighting. This is what you want for a data colour-map:
            PowerPoint's 3D renderer lights models with bright studio lights
            that otherwise clip a lit surface toward white and wash the colours
            out. Set False for a conventional lit/shaded matte surface.

    Raises:
        ValueError: if the colour count does not match faces/vertices.
    """
    V = np.asarray(vertices, dtype=np.float64)
    F = np.asarray(faces, dtype=np.int64)
    C = np.asarray(colors, dtype=np.float32)
    if C.ndim != 2 or C.shape[1] not in (3, 4):
        raise ValueError("colors must be (K,3) or (K,4)")

    if per_face:
        if C.shape[0] != F.shape[0]:
            raise ValueError(
                f"per-face colors ({C.shape[0]}) must match face count ({F.shape[0]})"
            )
        # Unweld: each face gets its own 3 vertices carrying the face colour.
        pos = V[F.reshape(-1)].astype(np.float32)
        col = np.repeat(C, 3, axis=0)
        idx = np.arange(F.shape[0] * 3, dtype=np.uint32)
    else:
        if C.shape[0] != V.shape[0]:
            raise ValueError(
                f"per-vertex colors ({C.shape[0]}) must match vertex count ({V.shape[0]})"
            )
        pos = V.astype(np.float32)
        col = C
        idx = F.reshape(-1).astype(np.uint32)

    # Promote RGB to RGBA (alpha = 1) so every glTF viewer accepts COLOR_0.
    if col.shape[1] == 3:
        col = np.concatenate(
            [col, np.ones((col.shape[0], 1), np.float32)], axis=1
        )
    col = np.ascontiguousarray(col, dtype=np.float32)
    pos = np.ascontiguousarray(pos, dtype=np.float32)

    pos_b = _pad4(pos.tobytes())
    col_b = _pad4(col.tobytes())
    idx_b = _pad4(idx.tobytes())
    buffer = pos_b + col_b + idx_b
    pos_off, col_off, idx_off = 0, len(pos_b), len(pos_b) + len(col_b)
    nvert = int(pos.shape[0])

    gltf = {
        "asset": {"version": "2.0", "generator": "ppt_image_inserter.glb"},
        "scene": 0,
        **({"extensionsUsed": ["KHR_materials_unlit"]} if unlit else {}),
        "scenes": [{"nodes": [0]}],
        "nodes": [{"mesh": 0}],
        "meshes": [{"primitives": [{
            "attributes": {"POSITION": 0, "COLOR_0": 1},
            "indices": 2,
            "mode": _TRIANGLES,
            "material": 0,
        }]}],
        # Matte, non-metallic material so the vertex colors render as diffuse
        # albedo (not glTF's default metallicFactor=1.0, which turns the surface
        # metallic and reflects scene light). With unlit=True we also tag
        # KHR_materials_unlit so PowerPoint's bright studio lights don't clip
        # the color toward white -- essential for a faithful data color-map.
        "materials": [_material(unlit)],
        "buffers": [{"byteLength": len(buffer)}],
        "bufferViews": [
            {"buffer": 0, "byteOffset": pos_off, "byteLength": len(pos_b), "target": _ARRAY_BUFFER},
            {"buffer": 0, "byteOffset": col_off, "byteLength": len(col_b), "target": _ARRAY_BUFFER},
            {"buffer": 0, "byteOffset": idx_off, "byteLength": len(idx_b), "target": _ELEMENT_ARRAY_BUFFER},
        ],
        "accessors": [
            {"bufferView": 0, "componentType": _FLOAT, "count": nvert, "type": "VEC3",
             "min": pos.min(axis=0).tolist(), "max": pos.max(axis=0).tolist()},
            {"bufferView": 1, "componentType": _FLOAT, "count": nvert, "type": "VEC4"},
            {"bufferView": 2, "componentType": _UINT, "count": int(idx.shape[0]), "type": "SCALAR"},
        ],
    }

    json_raw = json.dumps(gltf, separators=(",", ":")).encode("utf-8")
    json_b = json_raw + b" " * ((-len(json_raw)) % 4)  # JSON chunk pads w/ spaces
    bin_b = _pad4(buffer)
    total = 12 + 8 + len(json_b) + 8 + len(bin_b)
    with open(out_path, "wb") as fh:
        fh.write(b"glTF" + struct.pack("<II", 2, total))
        fh.write(struct.pack("<I", len(json_b)) + b"JSON" + json_b)
        fh.write(struct.pack("<I", len(bin_b)) + b"BIN\x00" + bin_b)


def read_glb_mesh(path):
    """Read a single-primitive .glb into (vertices, faces, colors).

    Assumes one mesh/primitive with float ``POSITION``, unsigned-int ``indices``
    and optional float ``COLOR_0`` (VEC3/VEC4) -- the shape produced by
    :func:`build_colored_glb`.

    Returns:
        (vertices (N,3) float, faces (M,3) int, colors (N,3) float in [0,1] or
        ``None`` if the mesh has no vertex colors).
    """
    with open(path, "rb") as fh:
        fh.read(4); struct.unpack("<II", fh.read(8))
        clen, _ = struct.unpack("<I4s", fh.read(8)); j = json.loads(fh.read(clen))
        blen, _ = struct.unpack("<I4s", fh.read(8)); buf = fh.read(blen)

    prim = j["meshes"][0]["primitives"][0]

    def accessor(idx, comp, dtype):
        a = j["accessors"][idx]
        bv = j["bufferViews"][a["bufferView"]]
        off = bv.get("byteOffset", 0) + a.get("byteOffset", 0)
        flat = np.frombuffer(buf, dtype=dtype, count=a["count"] * comp, offset=off)
        return flat.reshape(a["count"], comp) if comp > 1 else flat

    pos = accessor(prim["attributes"]["POSITION"], 3, "<f4").astype(np.float64)
    itype = j["accessors"][prim["indices"]]["componentType"]
    idx = accessor(prim["indices"], 1, {5125: "<u4", 5123: "<u2", 5121: "<u1"}[itype])
    faces = idx.reshape(-1, 3).astype(np.int64)

    colors = None
    if "COLOR_0" in prim["attributes"]:
        ci = prim["attributes"]["COLOR_0"]
        comp = 4 if j["accessors"][ci]["type"] == "VEC4" else 3
        colors = accessor(ci, comp, "<f4")[:, :3].astype(np.float64)
    return pos, faces, colors


def recolor_glb_for_powerpoint(in_path, out_path, saturate=1.0, darken=1.0):
    """Re-emit a colored .glb tuned for PowerPoint's 3D renderer.

    PowerPoint lights 3D models with bright studio lights that desaturate a
    lit surface, so a faithful data color-map comes out pale. This reads a
    colored ``.glb``, optionally boosts saturation and darkens the vertex
    colors to pre-compensate, and rewrites it with the matte/unlit material
    (:func:`build_colored_glb` with ``unlit=True``). With the defaults
    (``saturate=1.0, darken=1.0``) it only adds the material -- fixing the
    metallic-default wash-out without changing colors.

    Args:
        in_path: source colored ``.glb``.
        out_path: destination ``.glb``.
        saturate: per-vertex saturation multiplier (push away from grey).
            ``~2.0-2.5`` roughly restores vividness lost to PowerPoint lights.
        darken: overall brightness multiplier (``~0.8`` counters the lift).

    Raises:
        ValueError: if the source ``.glb`` has no vertex colors.
    """
    vertices, faces, colors = read_glb_mesh(in_path)
    if colors is None:
        raise ValueError(f"{in_path} has no COLOR_0 vertex colors to recolor")
    if saturate != 1.0 or darken != 1.0:
        luma = (colors * _LUMA).sum(axis=1, keepdims=True)
        colors = np.clip(luma + saturate * (colors - luma), 0.0, 1.0) * darken
    build_colored_glb(vertices, faces, colors, out_path, per_face=False, unlit=True)
