"""Movie-related helpers for PowerPoint decks.

Currently this module exposes a single helper that fixes the well-known
"video does not autoplay" issue with PowerPoint media inserted via either
``python-pptx``'s ``shapes.add_movie`` or PowerPoint COM's
``Shapes.AddMediaObject2``: both pipelines write a ``<p:cTn>`` start
condition of ``delay="indefinite"`` (click-to-start) on the trigger node
above the media-play effect, and PowerPoint surfaces that as click-to-start
in the Animation Pane and in normal mode regardless of what
``PlaySettings.PlayOnEntry`` reports via COM.

The fix is a small OOXML rewrite that PowerPoint then preserves on
subsequent saves.
"""

from __future__ import annotations

import os
import zipfile
from typing import Tuple

from lxml import etree


P_NAMESPACE = "http://schemas.openxmlformats.org/presentationml/2006/main"
_CTN_TAG = f"{{{P_NAMESPACE}}}cTn"
_STCONDLST_TAG = f"{{{P_NAMESPACE}}}stCondLst"
_COND_TAG = f"{{{P_NAMESPACE}}}cond"


def force_autoplay_in_pptx(pptx_path: str) -> int:
    """Rewrite click-to-start triggers on media-play effects to autoplay.

    Opens the .pptx as a zip, walks every ``ppt/slides/slideN.xml``, and for
    each ``<p:cTn>`` that wraps a ``presetClass="mediacall" presetID="1"``
    media-play effect rewrites the trigger's ``<p:cond delay="indefinite"/>``
    start condition to ``<p:cond delay="0"/>``. This is the canonical OOXML
    that PowerPoint UI generates for "Start: Automatically", and PowerPoint
    will not normalize it back on the next save.

    The fallback ``<p:cond evt="onBegin" ...>`` condition is left untouched.

    Args:
        pptx_path: Path to the .pptx file. The file is rewritten in place.

    Returns:
        The number of slides whose timing tree was modified.
    """
    tmp_path = pptx_path + ".autoplay.tmp"
    rewritten = 0
    with zipfile.ZipFile(pptx_path, "r") as zin, zipfile.ZipFile(
        tmp_path, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if _is_slide_xml(item.filename):
                new_data, changed = _patch_slide_xml(data)
                if changed:
                    rewritten += 1
                data = new_data
            zout.writestr(item, data)
    os.replace(tmp_path, pptx_path)
    return rewritten


def _is_slide_xml(name: str) -> bool:
    """Return True for ``ppt/slides/slideN.xml`` entries (not slideLayouts)."""
    if not name.startswith("ppt/slides/slide"):
        return False
    if not name.endswith(".xml"):
        return False
    return "/" not in name[len("ppt/slides/") :]


def _patch_slide_xml(xml_bytes: bytes) -> Tuple[bytes, bool]:
    """Return possibly-rewritten slide XML and a changed flag."""
    root = etree.fromstring(xml_bytes)
    changed = False

    for play_ctn in root.iter(_CTN_TAG):
        if play_ctn.get("presetClass") != "mediacall":
            continue
        if play_ctn.get("presetID") != "1":
            continue

        trigger_ctn = _find_trigger_ctn(play_ctn)
        if trigger_ctn is None:
            continue

        cond_list = trigger_ctn.find(_STCONDLST_TAG)
        if cond_list is None:
            continue

        for cond in cond_list.findall(_COND_TAG):
            if cond.get("evt") is not None:
                continue
            if cond.get("delay") != "indefinite":
                continue
            cond.set("delay", "0")
            changed = True

    if not changed:
        return xml_bytes, False

    out = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    return out, True


def _find_trigger_ctn(play_ctn: etree._Element):
    """Walk up the cTn chain to the trigger cTn directly under mainSeq.

    The OOXML structure is::

        <p:cTn id="2" nodeType="mainSeq">
          <p:childTnLst><p:par>
            <p:cTn id="3" ...>          # trigger (target of this walk)
              <p:childTnLst><p:par>
                <p:cTn id="4" ...>
                  <p:childTnLst><p:par>
                    <p:cTn id="5" presetClass="mediacall" .../>  # play_ctn

    Each step up walks one cTn level: cTn -> par -> childTnLst -> cTn.
    Stops when the next cTn ancestor is the mainSeq node.
    """
    current = play_ctn
    while current is not None:
        par = current.getparent()
        if par is None:
            return None
        child_tn_lst = par.getparent()
        if child_tn_lst is None:
            return None
        ancestor_ctn = child_tn_lst.getparent()
        if ancestor_ctn is None or ancestor_ctn.tag != _CTN_TAG:
            return None
        if ancestor_ctn.get("nodeType") == "mainSeq":
            return current
        current = ancestor_ctn
    return None
