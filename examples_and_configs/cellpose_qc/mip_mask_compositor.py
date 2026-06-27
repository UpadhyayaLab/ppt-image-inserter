"""Dataset-agnostic helpers for building (raw MIP | cellpose mask) composites.

All functions are pure: they take paths/arrays and return arrays/images or
write to an explicit output path. No hard-coded data paths, no global state.
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, Iterable, Optional, Tuple, Union

import numpy as np
import tifffile
from PIL import Image, ImageDraw, ImageFont

PathLike = Union[str, Path]


def load_actin_mip(raw_tif_path: PathLike, actin_channel: int) -> np.ndarray:
    """Read a multi-channel z-stack TIFF, return the XY MIP for one channel.

    Fast path: ImageJ hyperstacks (``is_imagej`` and metadata has
    ``channels``/``slices``). Page indexing is ``z * n_channels + c`` —
    only the actin channel's pages are read off disk, and the MIP is
    accumulated incrementally so peak memory stays at one plane.

    Fallback path: full ``tifffile.imread`` followed by axis inference,
    used for non-ImageJ 3-D / 4-D TIFFs.

    Returns a 2-D uint16 array (or whatever the source dtype is).
    """
    with tifffile.TiffFile(str(raw_tif_path)) as tf:
        if tf.is_imagej and tf.imagej_metadata:
            meta = tf.imagej_metadata
            n_pages = len(tf.pages)
            n_c = int(meta.get("channels", 1))
            n_z = int(meta.get("slices", n_pages // max(n_c, 1)))
            if not 0 <= actin_channel < n_c:
                raise IndexError(
                    f"actin_channel={actin_channel} out of range for "
                    f"channels={n_c} (file {raw_tif_path})."
                )
            mip: Optional[np.ndarray] = None
            for z in range(n_z):
                page_idx = z * n_c + actin_channel
                if page_idx >= n_pages:
                    break
                plane = tf.pages[page_idx].asarray()
                if mip is None:
                    mip = plane.copy()
                else:
                    np.maximum(mip, plane, out=mip)
            if mip is None:
                raise ValueError(f"No pages read from {raw_tif_path}")
            return mip

    # Non-ImageJ fallback: full read + axis inference (kept for
    # forward-compat with other dataset layouts).
    arr = tifffile.imread(str(raw_tif_path))
    if arr.ndim == 3:
        return arr.max(axis=0)
    if arr.ndim != 4:
        raise ValueError(
            f"Unsupported TIFF shape {arr.shape} for {raw_tif_path}; "
            "expected 3-D (Z,Y,X) or 4-D (Z,C,Y,X) / (C,Z,Y,X)."
        )

    a0, a1 = arr.shape[0], arr.shape[1]
    if a0 == a1:
        raise ValueError(
            f"Cannot infer channel axis: leading axes are equal in {arr.shape}."
        )
    channel_axis = 0 if a0 < a1 else 1
    if not 0 <= actin_channel < arr.shape[channel_axis]:
        raise IndexError(
            f"actin_channel={actin_channel} out of range for axis size "
            f"{arr.shape[channel_axis]} (file {raw_tif_path})."
        )
    if channel_axis == 1:
        plane_stack = arr[:, actin_channel, :, :]
    else:
        plane_stack = arr[actin_channel, :, :, :]
    return plane_stack.max(axis=0)


def load_channel_mips(
    raw_tif_path: PathLike, channel_indices: Iterable[int]
) -> Dict[int, np.ndarray]:
    """Read one or more channel MIPs from an ImageJ hyperstack in one pass.

    Walks the TIFF's pages once and accumulates a max-projection for every
    requested channel; per-channel peak memory is one plane. Returns a
    dict ``{channel_index: 2-D MIP}``.

    Non-ImageJ files fall back to full-volume reads (one read per channel
    via the existing single-channel function).
    """
    indices = list(dict.fromkeys(int(c) for c in channel_indices))  # uniq, ordered
    if not indices:
        raise ValueError("channel_indices must contain at least one channel")

    with tifffile.TiffFile(str(raw_tif_path)) as tf:
        if tf.is_imagej and tf.imagej_metadata:
            meta = tf.imagej_metadata
            n_pages = len(tf.pages)
            n_c = int(meta.get("channels", 1))
            n_z = int(meta.get("slices", n_pages // max(n_c, 1)))
            for c in indices:
                if not 0 <= c < n_c:
                    raise IndexError(
                        f"channel {c} out of range for channels={n_c} "
                        f"(file {raw_tif_path})."
                    )
            mips: Dict[int, Optional[np.ndarray]] = {c: None for c in indices}
            for z in range(n_z):
                for c in indices:
                    page_idx = z * n_c + c
                    if page_idx >= n_pages:
                        continue
                    plane = tf.pages[page_idx].asarray()
                    if mips[c] is None:
                        mips[c] = plane.copy()
                    else:
                        np.maximum(mips[c], plane, out=mips[c])
            out: Dict[int, np.ndarray] = {}
            for c, m in mips.items():
                if m is None:
                    raise ValueError(
                        f"No pages read for channel {c} from {raw_tif_path}"
                    )
                out[c] = m
            return out

    # Non-ImageJ fallback: one full read per channel via load_actin_mip.
    return {c: load_actin_mip(raw_tif_path, c) for c in indices}


def make_two_color_rgb(
    gray_u8: np.ndarray,
    color_u8: np.ndarray,
    color_rgb: Tuple[int, int, int] = (0, 255, 255),
) -> Image.Image:
    """Combine a grayscale channel with a colorized second channel into RGB.

    The gray channel goes into all three RGB planes; the color channel is
    added scaled by ``color_rgb / 255``. Output saturates at 255.
    """
    if gray_u8.shape != color_u8.shape:
        raise ValueError(
            f"gray and color must have same shape; got {gray_u8.shape} vs "
            f"{color_u8.shape}"
        )
    g = gray_u8.astype(np.uint16)
    c = color_u8.astype(np.uint16)
    cr, cg, cb = color_rgb
    r = np.clip(g + (c * cr) // 255, 0, 255).astype(np.uint8)
    gg = np.clip(g + (c * cg) // 255, 0, 255).astype(np.uint8)
    b = np.clip(g + (c * cb) // 255, 0, 255).astype(np.uint8)
    rgb = np.stack([r, gg, b], axis=-1)
    return Image.fromarray(rgb, mode="RGB")


def labels_from_colored_rgb(rgb: np.ndarray) -> np.ndarray:
    """Re-derive integer labels from a colored mask RGB image.

    Each unique non-black RGB triplet becomes one label. Useful when the
    saved cellpose ``.npy`` / ``.tif`` is a colored visualization
    (`Actin_Crop.py` does this) rather than the raw integer labels.
    """
    if rgb.ndim != 3 or rgb.shape[-1] < 3:
        raise ValueError(f"Expected (Y,X,3) RGB array, got {rgb.shape}")
    flat = rgb[..., :3].astype(np.uint32)
    keys = (flat[..., 0] << 16) | (flat[..., 1] << 8) | flat[..., 2]
    labels = np.zeros(keys.shape, dtype=np.int32)
    unique = np.unique(keys)
    next_id = 1
    for k in unique:
        if k == 0:  # black = background
            continue
        labels[keys == k] = next_id
        next_id += 1
    return labels


def overlay_label_outlines(
    base_rgb: Image.Image,
    labels: np.ndarray,
    thickness_px: int = 2,
) -> Image.Image:
    """Paint cellpose label boundaries on a base RGB image, color per label.

    Uses ``skimage.segmentation.find_boundaries`` (mode='outer'); each label
    id picks a color from ``matplotlib.cm.tab20`` (cycled). Boundary lines
    are dilated to ``thickness_px`` for slide visibility.
    """
    from skimage.segmentation import find_boundaries
    from skimage.morphology import binary_dilation, disk
    import matplotlib.cm as cm

    if labels.ndim != 2:
        raise ValueError(f"labels must be 2-D, got shape {labels.shape}")

    base = base_rgb.convert("RGB")
    if base.size != (labels.shape[1], labels.shape[0]):
        base = base.resize((labels.shape[1], labels.shape[0]), Image.NEAREST)
    out = np.array(base, dtype=np.uint8)

    boundaries = find_boundaries(labels, mode="outer")
    if thickness_px > 1:
        boundaries = binary_dilation(boundaries, footprint=disk(max(1, thickness_px // 2)))

    boundary_labels = np.where(boundaries, labels, 0)
    unique = np.unique(boundary_labels)
    unique = unique[unique > 0]
    if unique.size == 0:
        return Image.fromarray(out, mode="RGB")

    cmap = cm.get_cmap("tab20", 20)
    palette = (np.array(cmap.colors)[:, :3] * 255).astype(np.uint8)
    for lbl in unique:
        color = palette[(int(lbl) - 1) % 20]
        m = boundary_labels == lbl
        out[m] = color
    return Image.fromarray(out, mode="RGB")


def to_display_uint8(
    arr: np.ndarray,
    low_pct: float = 1.0,
    high_pct: float = 99.5,
    min_lo: Optional[float] = None,
) -> np.ndarray:
    """Percentile contrast-stretch a 2-D array to uint8 for display.

    ``min_lo`` (in raw input units) floors the lower clip — useful when a
    channel is dim and the percentile-derived ``lo`` is below background,
    which would otherwise leak background into the displayed image (e.g.
    a faint red haze when overlaying Hoechst on actin).
    """
    if arr.ndim != 2:
        raise ValueError(f"Expected 2-D array, got shape {arr.shape}")
    a = arr.astype(np.float32)
    lo, hi = np.percentile(a, [low_pct, high_pct])
    if min_lo is not None:
        lo = max(float(lo), float(min_lo))
    if hi <= lo:
        hi = lo + 1.0
    a = np.clip((a - lo) / (hi - lo), 0.0, 1.0)
    return (a * 255).astype(np.uint8)


def load_mask_rgb(mask_tif_path: PathLike) -> Image.Image:
    """Load a cellpose mask TIFF as an RGB PIL image.

    If the file is already an RGB visualization (Y, X, 3), use it directly.
    If it's a 2-D integer label image, color the labels with a colormap so
    the deck is still useful when only raw label outputs exist.
    """
    arr = tifffile.imread(str(mask_tif_path))
    if arr.ndim == 3 and arr.shape[-1] in (3, 4):
        rgb = arr[..., :3].astype(np.uint8)
        return Image.fromarray(rgb, mode="RGB")
    if arr.ndim == 2:
        return _colorize_labels(arr)
    raise ValueError(
        f"Unsupported mask shape {arr.shape} for {mask_tif_path}; "
        "expected (Y,X,3) RGB or (Y,X) labels."
    )


def _colorize_labels(labels: np.ndarray) -> Image.Image:
    """Map a 2-D integer label image to RGB using tab20 cycled by label id."""
    import matplotlib.cm as cm  # lazy import; only needed in fallback

    labels = labels.astype(np.int64)
    n_colors = 20
    cmap = cm.get_cmap("tab20", n_colors)
    rgb = np.zeros(labels.shape + (3,), dtype=np.uint8)
    mask = labels > 0
    if mask.any():
        color_idx = ((labels[mask] - 1) % n_colors)
        colors = (np.array(cmap.colors)[color_idx][:, :3] * 255).astype(np.uint8)
        rgb[mask] = colors
    return Image.fromarray(rgb, mode="RGB")


def add_scalebar(
    img: Image.Image,
    length_um: float,
    pixel_size_um: float,
    margin_px: int = 50,
    bar_thickness_px: int = 18,
    label: Optional[str] = None,
    color: Tuple[int, int, int] = (255, 255, 255),
    font_size_px: int = 64,
) -> Image.Image:
    """Draw a horizontal scalebar in the upper-left of ``img``.

    Returns a NEW RGB image; ``img`` is not mutated. ``pixel_size_um`` is
    the size of one pixel of ``img`` in microns at the moment of drawing —
    if the image is resized later, the bar resizes proportionally and the
    physical length is preserved.
    """
    if pixel_size_um <= 0:
        raise ValueError(f"pixel_size_um must be > 0 (got {pixel_size_um})")
    canvas = img.convert("RGB").copy()
    draw = ImageDraw.Draw(canvas)

    bar_w = max(1, int(round(length_um / pixel_size_um)))
    x0 = margin_px
    y0 = margin_px
    draw.rectangle([x0, y0, x0 + bar_w, y0 + bar_thickness_px], fill=color)

    text = label if label is not None else f"{length_um:g} μm"
    font = _load_font(font_size_px)
    text_y = y0 + bar_thickness_px + 8
    draw.text((x0, text_y), text, fill=color, font=font)
    return canvas


def _load_font(size_px: int) -> ImageFont.ImageFont:
    """Best-effort TrueType font load with default fallback."""
    for name in ("arial.ttf", "DejaVuSans.ttf", "segoeui.ttf"):
        try:
            return ImageFont.truetype(name, size=size_px)
        except OSError:
            continue
    return ImageFont.load_default()


def compose_side_by_side(
    left_img: Image.Image,
    right_img: Image.Image,
    gap_px: int = 4,
    bg: Tuple[int, int, int] = (0, 0, 0),
    divider_color: Optional[Tuple[int, int, int]] = (255, 255, 255),
) -> Image.Image:
    """Paste two images side by side at matched height with a divider between.

    When ``divider_color`` is not None, the inter-image gap is filled with
    that color (acts as a divider line). Pass ``divider_color=None`` for a
    plain ``bg``-colored gap.
    """
    target_h = min(left_img.height, right_img.height)
    left = _resize_to_height(left_img, target_h)
    right = _resize_to_height(right_img, target_h)
    total_w = left.width + gap_px + right.width
    canvas = Image.new("RGB", (total_w, target_h), bg)
    canvas.paste(left.convert("RGB"), (0, 0))
    canvas.paste(right.convert("RGB"), (left.width + gap_px, 0))
    if divider_color is not None and gap_px > 0:
        ImageDraw.Draw(canvas).rectangle(
            [left.width, 0, left.width + gap_px - 1, target_h - 1],
            fill=divider_color,
        )
    return canvas


def _resize_to_height(img: Image.Image, target_h: int) -> Image.Image:
    if img.height == target_h:
        return img
    new_w = max(1, round(img.width * target_h / img.height))
    return img.resize((new_w, target_h), Image.LANCZOS)


def build_composite(
    raw_tif_path: PathLike,
    mask_tif_path: PathLike,
    actin_channel: int,
    out_path: PathLike,
    pixel_size_um: float,
    scalebar_um: float = 10.0,
    contrast_low_pct: float = 1.0,
    contrast_high_pct: float = 99.5,
    gap_px: int = 4,
    image_format: str = "png",
    jpeg_quality: int = 90,
) -> Path:
    """End-to-end: read raw + mask, draw scalebar, build composite, write image.

    ``pixel_size_um`` is required and applied regardless of TIFF metadata
    (metadata is often wrong for ND2-converted hyperstacks).

    ``image_format`` is ``"png"`` (lossless) or ``"jpg"``/``"jpeg"`` (smaller
    file at the cost of compression artifacts; ``jpeg_quality`` 1-95).
    """
    mip = load_actin_mip(raw_tif_path, actin_channel)
    mip_u8 = to_display_uint8(mip, contrast_low_pct, contrast_high_pct)
    left = Image.fromarray(mip_u8, mode="L").convert("RGB")
    left = add_scalebar(left, length_um=scalebar_um, pixel_size_um=pixel_size_um)
    right = load_mask_rgb(mask_tif_path)
    composite = compose_side_by_side(left, right, gap_px=gap_px)

    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    fmt = image_format.lower()
    if fmt in ("jpg", "jpeg"):
        composite.save(out, format="JPEG", quality=int(jpeg_quality), optimize=True)
    elif fmt == "png":
        composite.save(out, format="PNG")
    else:
        raise ValueError(f"image_format must be 'png' or 'jpg' (got {image_format!r})")
    return out


# Backwards-compat alias: older callers used build_composite_png.
build_composite_png = build_composite


def _save_composite(
    img: Image.Image, out_path: Path, image_format: str, jpeg_quality: int
) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fmt = image_format.lower()
    if fmt in ("jpg", "jpeg"):
        img.save(out_path, format="JPEG", quality=int(jpeg_quality), optimize=True)
    elif fmt == "png":
        img.save(out_path, format="PNG")
    else:
        raise ValueError(f"image_format must be 'png' or 'jpg' (got {image_format!r})")


def build_fov_composites(
    raw_tif_path: PathLike,
    mask_tif_path: PathLike,
    actin_channel: int,
    out_plain: Optional[PathLike],
    out_with_nuc: Optional[PathLike],
    pixel_size_um: float,
    scalebar_um: float = 10.0,
    actin_low_pct: float = 1.0,
    actin_high_pct: float = 99.5,
    hoechst_channel: Optional[int] = None,
    hoechst_low_pct: float = 2.0,
    hoechst_high_pct: float = 99.5,
    hoechst_min_lo: Optional[float] = 120.0,
    hoechst_color: Tuple[int, int, int] = (255, 0, 0),
    gap_px: int = 4,
    image_format: str = "jpg",
    jpeg_quality: int = 90,
) -> list:
    """Read raw + mask once, write up to two composites per FOV.

    - ``out_plain``: write actin + colored mask (the v1 composite).
    - ``out_with_nuc``: also write actin + Hoechst overlay + colored mask
      (requires ``hoechst_channel`` to be set).
    Either path may be ``None`` to skip that output. Returns the list of
    paths actually written. Skipping reduces I/O: the with-nuc variant
    triggers a Hoechst-page read only when its output is needed.
    """
    want_plain = out_plain is not None
    want_with_nuc = out_with_nuc is not None and hoechst_channel is not None
    if not want_plain and not want_with_nuc:
        return []

    channels = [actin_channel]
    if want_with_nuc:
        channels.append(int(hoechst_channel))
    mips = load_channel_mips(raw_tif_path, channels)
    actin_u8 = to_display_uint8(mips[actin_channel], actin_low_pct, actin_high_pct)

    # Right panel shared by both composites.
    right = load_mask_rgb(mask_tif_path)

    written = []

    if want_plain:
        actin_rgb = Image.fromarray(actin_u8, mode="L").convert("RGB")
        left = add_scalebar(actin_rgb, length_um=scalebar_um, pixel_size_um=pixel_size_um)
        composite = compose_side_by_side(left, right, gap_px=gap_px)
        _save_composite(composite, Path(out_plain), image_format, jpeg_quality)
        written.append(Path(out_plain))

    if want_with_nuc:
        hoechst_u8 = to_display_uint8(
            mips[int(hoechst_channel)], hoechst_low_pct, hoechst_high_pct,
            min_lo=hoechst_min_lo,
        )
        left2 = make_two_color_rgb(actin_u8, hoechst_u8, color_rgb=hoechst_color)
        left2 = add_scalebar(left2, length_um=scalebar_um, pixel_size_um=pixel_size_um)
        composite2 = compose_side_by_side(left2, right, gap_px=gap_px)
        _save_composite(composite2, Path(out_with_nuc), image_format, jpeg_quality)
        written.append(Path(out_with_nuc))

    return written


def build_composite_v2(
    raw_tif_path: PathLike,
    mask_tif_path: PathLike,
    actin_channel: int,
    hoechst_channel: int,
    out_path: PathLike,
    pixel_size_um: float,
    scalebar_um: float = 10.0,
    actin_low_pct: float = 1.0,
    actin_high_pct: float = 99.5,
    hoechst_low_pct: float = 2.0,
    hoechst_high_pct: float = 99.5,
    hoechst_min_lo: Optional[float] = 120.0,
    hoechst_color: Tuple[int, int, int] = (0, 255, 255),
    gap_px: int = 4,
    image_format: str = "jpg",
    jpeg_quality: int = 90,
) -> Path:
    """End-to-end v2: actin+Hoechst overlay (left) + colored cellpose mask (right).

    Left panel: actin (gray) blended with Hoechst (in ``hoechst_color``),
    plus a 10 μm scalebar. ``hoechst_min_lo`` floors the Hoechst lower
    clip so dim FOVs don't bleed background red.
    Right panel: the cellpose colored mask RGB (same as v1).
    """
    mips = load_channel_mips(raw_tif_path, [actin_channel, hoechst_channel])
    actin_u8 = to_display_uint8(mips[actin_channel], actin_low_pct, actin_high_pct)
    hoechst_u8 = to_display_uint8(
        mips[hoechst_channel], hoechst_low_pct, hoechst_high_pct,
        min_lo=hoechst_min_lo,
    )

    left = make_two_color_rgb(actin_u8, hoechst_u8, color_rgb=hoechst_color)
    left = add_scalebar(left, length_um=scalebar_um, pixel_size_um=pixel_size_um)

    right = load_mask_rgb(mask_tif_path)

    composite = compose_side_by_side(left, right, gap_px=gap_px)

    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    fmt = image_format.lower()
    if fmt in ("jpg", "jpeg"):
        composite.save(out, format="JPEG", quality=int(jpeg_quality), optimize=True)
    elif fmt == "png":
        composite.save(out, format="PNG")
    else:
        raise ValueError(f"image_format must be 'png' or 'jpg' (got {image_format!r})")
    return out
