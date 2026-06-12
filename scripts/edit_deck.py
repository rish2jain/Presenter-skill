#!/usr/bin/env python3
"""
edit_deck.py — Surgical editing of existing .pptx files via OOXML unpack/pack.

Subcommands:
    unpack <deck.pptx> <dir>          Extract + pretty-print XML for editing
    pack <dir> <out.pptx>             Validate XML and re-zip into a .pptx
    list <dir>                        Show slide order, files, and titles
    duplicate <dir> <N>               Duplicate slide at deck position N (see list)
    remove <dir> <N>                  Remove slide at deck position N
    reorder <dir> <order>             Reorder slides: e.g. 3,1,2,4 (positions from list)
    clean <dir>                       Remove unreferenced media from ppt/media/
    extract <deck.pptx> <sel> --output sub.pptx
                                      Split: new deck from positions "1,3-5"
    append <dst.pptx> <src.pptx> [--slides 2-4] --output merged.pptx
                                      Merge: copy src slides after dst's last

Workflow: unpack → (structural ops: duplicate/remove) → edit slide XML text →
pack → QA with qa_check.py + render_slides.py. See references/editing.md.
"""
import argparse
import posixpath
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
SLIDE_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
NOTES_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
LAYOUT_RT = ("http://schemas.openxmlformats.org/officeDocument/2006/"
             "relationships/slideLayout")
MASTER_RT = ("http://schemas.openxmlformats.org/officeDocument/2006/"
             "relationships/slideMaster")


def _xml(path):
    return etree.parse(str(path))


def _write_xml(tree, path):
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)


# ── unpack / pack ────────────────────────────────────────────────────────────
def unpack(pptx, out_dir):
    out = Path(out_dir)
    if out.exists():
        shutil.rmtree(out)
    with zipfile.ZipFile(pptx) as zf:
        zf.extractall(out)
    # Pretty-print XML so the Edit tool can target lines. lxml only indents
    # element-only content, so runs/text nodes are untouched.
    count = 0
    for f in out.rglob("*"):
        if f.suffix.lower() in (".xml", ".rels") and f.is_file():
            try:
                tree = etree.parse(str(f))
                f.write_bytes(etree.tostring(tree, pretty_print=True,
                                             xml_declaration=True,
                                             encoding="UTF-8", standalone=True))
                count += 1
            except etree.XMLSyntaxError:
                pass
    print(f"Unpacked {pptx} → {out}/ ({count} XML files formatted)")
    print("Slides:", ", ".join(s.name for s in sorted(
        (out / "ppt" / "slides").glob("slide*.xml"),
        key=lambda x: int(re.search(r"\d+", x.name).group()))))


def pack(src_dir, out_pptx):
    src = Path(src_dir)
    bad = []
    for f in src.rglob("*"):
        if f.suffix.lower() in (".xml", ".rels") and f.is_file():
            try:
                etree.parse(str(f))
            except etree.XMLSyntaxError as e:
                bad.append(f"{f.relative_to(src)}: {e}")
    if bad:
        print("XML validation failed — fix before packing:", file=sys.stderr)
        for b in bad:
            print(f"  [ERROR] {b}", file=sys.stderr)
        return False

    files = sorted(p for p in src.rglob("*") if p.is_file())
    # [Content_Types].xml must be present; put it first in the archive
    files.sort(key=lambda p: p.name != "[Content_Types].xml")
    with zipfile.ZipFile(out_pptx, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(f, f.relative_to(src).as_posix())
    print(f"Packed {src}/ → {out_pptx} ({len(files)} files)")
    return True


# ── slide structure helpers ──────────────────────────────────────────────────
def _presentation_parts(src):
    src = Path(src)
    return (src / "ppt" / "presentation.xml",
            src / "ppt" / "_rels" / "presentation.xml.rels",
            src / "[Content_Types].xml")


def _slide_entries(pres_tree, rels_tree):
    """Ordered [(sldId el, rId, slide filename)] from presentation.xml."""
    rid_to_target = {
        r.get("Id"): r.get("Target")
        for r in rels_tree.getroot().findall(f"{{{NS['rel']}}}Relationship")
        if r.get("Type") == SLIDE_RT
    }
    entries = []
    for sld in pres_tree.getroot().findall(f"p:sldIdLst/p:sldId", NS):
        rid = sld.get(f"{{{NS['r']}}}id")
        entries.append((sld, rid, Path(rid_to_target[rid]).name))
    return entries


def _slide_file_num(fname):
    """'slide3.xml' -> 3."""
    m = re.search(r"\d+", fname)
    return int(m.group()) if m else None


def _resolve_position(src_dir, position):
    """Map 1-based deck position (as shown by `list`) to slide file number."""
    pres_p, rels_p, _ = _presentation_parts(src_dir)
    entries = _slide_entries(_xml(pres_p), _xml(rels_p))
    if position < 1 or position > len(entries):
        print(f"ERROR: position {position} out of range (1-{len(entries)})",
              file=sys.stderr)
        return None
    fname = entries[position - 1][2]
    num = _slide_file_num(fname)
    if num is None:
        print(f"ERROR: cannot parse slide number from {fname}", file=sys.stderr)
    return num


def list_slides(src_dir):
    pres_p, rels_p, _ = _presentation_parts(src_dir)
    pres, rels = _xml(pres_p), _xml(rels_p)
    for pos, (_, rid, fname) in enumerate(_slide_entries(pres, rels), 1):
        slide_path = Path(src_dir) / "ppt" / "slides" / fname
        texts = etree.parse(str(slide_path)).getroot().findall(f".//a:t", NS)
        first = next((t.text for t in texts if t.text and t.text.strip()), "")
        print(f"  {pos:2d}. {fname:<14} ({rid})  {first[:60]}")


def duplicate(src_dir, position):
    slide_num = _resolve_position(src_dir, position)
    if slide_num is None:
        return False
    src = Path(src_dir)
    slides_dir = src / "ppt" / "slides"
    src_xml = slides_dir / f"slide{slide_num}.xml"
    if not src_xml.exists():
        print(f"ERROR: {src_xml} not found", file=sys.stderr)
        return False

    nums = [int(re.search(r"\d+", f.name).group())
            for f in slides_dir.glob("slide*.xml")]
    new_num = max(nums) + 1
    new_xml = slides_dir / f"slide{new_num}.xml"
    shutil.copy(src_xml, new_xml)

    # slide rels: copy but drop the notesSlide relationship (notes aren't cloned)
    src_rels = slides_dir / "_rels" / f"slide{slide_num}.xml.rels"
    if src_rels.exists():
        rels = _xml(src_rels)
        root = rels.getroot()
        for r in root.findall(f"{{{NS['rel']}}}Relationship"):
            if r.get("Type") == NOTES_RT:
                root.remove(r)
        _write_xml(rels, slides_dir / "_rels" / f"slide{new_num}.xml.rels")

    # [Content_Types].xml override
    pres_p, rels_p, ct_p = _presentation_parts(src)
    ct = _xml(ct_p)
    override = etree.SubElement(ct.getroot(), f"{{{NS['ct']}}}Override")
    override.set("PartName", f"/ppt/slides/slide{new_num}.xml")
    override.set("ContentType",
                 "application/vnd.openxmlformats-officedocument"
                 ".presentationml.slide+xml")
    _write_xml(ct, ct_p)

    # presentation rels + sldIdLst (insert right after the source slide)
    pres, rels = _xml(pres_p), _xml(rels_p)
    rel_root = rels.getroot()
    max_rid = max(int(r.get("Id")[3:]) for r in
                  rel_root.findall(f"{{{NS['rel']}}}Relationship"))
    new_rid = f"rId{max_rid + 1}"
    rel = etree.SubElement(rel_root, f"{{{NS['rel']}}}Relationship")
    rel.set("Id", new_rid)
    rel.set("Type", SLIDE_RT)
    rel.set("Target", f"slides/slide{new_num}.xml")
    _write_xml(rels, rels_p)

    entries = _slide_entries(pres, _xml(rels_p))
    max_sldid = max(int(e[0].get("id")) for e in entries)
    source_el = next(e[0] for e in entries if e[2] == f"slide{slide_num}.xml")
    new_sld = etree.SubElement(source_el.getparent(), f"{{{NS['p']}}}sldId")
    new_sld.set("id", str(max_sldid + 1))
    new_sld.set(f"{{{NS['r']}}}id", new_rid)
    source_el.addnext(new_sld)
    _write_xml(pres, pres_p)

    print(f"Duplicated position {position} (slide{slide_num}.xml) → slide{new_num}.xml "
          f"(inserted after position {position})")
    return True


def remove(src_dir, position):
    slide_num = _resolve_position(src_dir, position)
    if slide_num is None:
        return False
    src = Path(src_dir)
    pres_p, rels_p, ct_p = _presentation_parts(src)
    pres, rels, ct = _xml(pres_p), _xml(rels_p), _xml(ct_p)

    target_fname = f"slide{slide_num}.xml"
    entries = _slide_entries(pres, rels)
    entry = next((e for e in entries if e[2] == target_fname), None)
    if entry is None:
        print(f"ERROR: {target_fname} is not in the slide list", file=sys.stderr)
        return False
    sld_el, rid, _ = entry

    sld_el.getparent().remove(sld_el)
    _write_xml(pres, pres_p)
    rel_root = rels.getroot()
    for r in rel_root.findall(f"{{{NS['rel']}}}Relationship"):
        if r.get("Id") == rid:
            rel_root.remove(r)
    _write_xml(rels, rels_p)
    ct_root = ct.getroot()
    for o in ct_root.findall(f"{{{NS['ct']}}}Override"):
        if o.get("PartName") == f"/ppt/slides/{target_fname}":
            ct_root.remove(o)
    _write_xml(ct, ct_p)

    for f in (src / "ppt" / "slides" / target_fname,
              src / "ppt" / "slides" / "_rels" / f"{target_fname}.rels"):
        f.unlink(missing_ok=True)
    print(f"Removed position {position} ({target_fname}; media orphans left in ppt/media/)")
    return True


def reorder(src_dir, order_str):
    """Reorder slides by 1-based positions from `list` output."""
    src = Path(src_dir)
    pres_p, rels_p, _ = _presentation_parts(src)
    pres, rels = _xml(pres_p), _xml(rels_p)
    entries = _slide_entries(pres, rels)
    n = len(entries)
    try:
        order = [int(x.strip()) for x in order_str.split(",")]
    except ValueError:
        print("ERROR: order must be comma-separated positions, e.g. 3,1,2,4",
              file=sys.stderr)
        return False
    if sorted(order) != list(range(1, n + 1)):
        print(f"ERROR: order must be a permutation of 1-{n}, got {order}",
              file=sys.stderr)
        return False

    lst = pres.getroot().find("p:sldIdLst", NS)
    children = [entries[i - 1][0] for i in order]
    for child in list(lst):
        lst.remove(child)
    for child in children:
        lst.append(child)
    _write_xml(pres, pres_p)
    print(f"Reordered slides to positions: {order_str}")
    return True


def clean_orphans(src_dir):
    """Delete ppt/media files not referenced by any .rels."""
    src = Path(src_dir)
    media = src / "ppt" / "media"
    if not media.is_dir():
        print("No ppt/media/ directory — nothing to clean.")
        return True
    referenced = set()
    for rels in src.rglob("*.rels"):
        text = rels.read_text(encoding="utf-8", errors="ignore")
        for m in re.finditer(r'(?:Target="[^"]*/)?media/([^"]+)"', text):
            referenced.add(m.group(1))
    removed = 0
    for f in media.iterdir():
        if f.is_file() and f.name not in referenced:
            f.unlink()
            removed += 1
            print(f"  Removed orphan: {f.name}")
    print(f"Clean complete: {removed} orphan file(s) removed.")
    return True


# ── extract / append (split & merge) ────────────────────────────────────────
def _parse_selection(spec, n):
    """'1,3-5' → [1, 3, 4, 5]. Ordered, deduped, validated against 1..n.
    Raises ValueError on empty/malformed/out-of-range/reversed input."""
    out, seen = [], set()
    for tok in (spec or "").split(","):
        tok = tok.strip()
        if not tok:
            continue
        m = re.fullmatch(r"(\d+)(?:-(\d+))?", tok)
        if not m:
            raise ValueError(f"bad selection token {tok!r} "
                             "(use positions like 2, 5-7)")
        a, b = int(m.group(1)), int(m.group(2) or m.group(1))
        if a > b:
            raise ValueError(f"reversed range {tok!r}")
        for p in range(a, b + 1):
            if p < 1 or p > n:
                raise ValueError(f"position {p} out of range (1-{n})")
            if p not in seen:
                seen.add(p)
                out.append(p)
    if not out:
        raise ValueError("empty selection")
    return out


def _part_rels_path(root, part):
    """ppt/slides/slide3.xml → <root>/ppt/slides/_rels/slide3.xml.rels."""
    p = Path(part)
    return Path(root) / p.parent / "_rels" / (p.name + ".rels")


def _resolve_target(base_part, target):
    """Relationship Target (relative to base part's dir) → package path."""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(
        posixpath.join(posixpath.dirname(base_part), target))


def _relative_target(from_part, to_part):
    return posixpath.relpath(to_part, posixpath.dirname(from_part))


def _iter_rels(tree):
    return tree.getroot().findall(f"{{{NS['rel']}}}Relationship")


def _gc_unreferenced(root):
    """Delete parts unreachable from the package roots; prune CT overrides.

    Used by extract: removing slides orphans their notesSlides/charts/
    embeddings/media — dangling parts invite PowerPoint repair prompts.
    """
    root = Path(root)
    reachable = set()

    def walk(part):
        if part in reachable:
            return
        reachable.add(part)
        rels = _part_rels_path(root, part)
        if rels.exists():
            for rel in _iter_rels(_xml(rels)):
                if rel.get("TargetMode") == "External":
                    continue
                walk(_resolve_target(part, rel.get("Target")))

    for rel in _iter_rels(_xml(root / "_rels" / ".rels")):
        if rel.get("TargetMode") != "External":
            walk(_resolve_target("", rel.get("Target")))

    removed = []
    for f in sorted(root.rglob("*")):
        if not f.is_file():
            continue
        part = f.relative_to(root).as_posix()
        if part == "[Content_Types].xml" or part == "_rels/.rels":
            continue
        if f.parent.name == "_rels":
            owner = (Path(part).parent.parent / f.name[:-len(".rels")]).as_posix()
            keep = owner in reachable
        else:
            keep = part in reachable
        if not keep:
            f.unlink()
            removed.append(part)

    if removed:
        ct_p = root / "[Content_Types].xml"
        ct = _xml(ct_p)
        gone = {f"/{p}" for p in removed}
        for o in ct.getroot().findall(f"{{{NS['ct']}}}Override"):
            if o.get("PartName") in gone:
                ct.getroot().remove(o)
        _write_xml(ct, ct_p)
        print(f"  GC: removed {len(removed)} unreferenced part(s)")


def extract(pptx, selection, output):
    """Split: copy slides at the selected 1-based positions into a new deck."""
    pptx, output = Path(pptx), Path(output)
    if output.resolve() == pptx.resolve():
        print("ERROR: --output must differ from the input deck",
              file=sys.stderr)
        return False
    with tempfile.TemporaryDirectory() as td:
        work = Path(td) / "unpacked"
        unpack(pptx, work)
        pres_p, rels_p, _ = _presentation_parts(work)
        n = len(_slide_entries(_xml(pres_p), _xml(rels_p)))
        try:
            keep = set(_parse_selection(selection, n))
        except ValueError as e:
            print(f"ERROR: {e}", file=sys.stderr)
            return False
        for pos in sorted(set(range(1, n + 1)) - keep, reverse=True):
            if not remove(work, pos):
                return False
        _strip_dangling_slide_links(work)
        _gc_unreferenced(work)
        if not pack(work, output):
            return False
    print(f"Extracted {len(keep)} slide(s) from {pptx} → {output}")
    return True


def _ct_index(root):
    """[Content_Types].xml → ({'/part': ct}, {'ext': ct})."""
    tree = _xml(Path(root) / "[Content_Types].xml")
    overrides = {o.get("PartName"): o.get("ContentType")
                 for o in tree.getroot().findall(f"{{{NS['ct']}}}Override")}
    defaults = {d.get("Extension").lower(): d.get("ContentType")
                for d in tree.getroot().findall(f"{{{NS['ct']}}}Default")}
    return overrides, defaults


def _ct_register(dst_ct, src_ct, src_part, new_part):
    """Mirror src's content-type declaration for new_part into dst's tree."""
    overrides, defaults = src_ct
    dst_root = dst_ct.getroot()
    ct = overrides.get(f"/{src_part}")
    if ct is not None:
        if not any(o.get("PartName") == f"/{new_part}" for o in
                   dst_root.findall(f"{{{NS['ct']}}}Override")):
            o = etree.SubElement(dst_root, f"{{{NS['ct']}}}Override")
            o.set("PartName", f"/{new_part}")
            o.set("ContentType", ct)
        return
    ext = src_part.rsplit(".", 1)[-1].lower()
    ct = defaults.get(ext)
    if ct is None:
        raise ValueError(f"no content type for {src_part} in source deck")
    existing = {d.get("Extension").lower(): d.get("ContentType")
                for d in dst_root.findall(f"{{{NS['ct']}}}Default")}
    if ext not in existing:
        d = etree.SubElement(dst_root, f"{{{NS['ct']}}}Default")
        d.set("Extension", ext)
        d.set("ContentType", ct)
    elif existing[ext] != ct:  # rare: same extension, different type
        o = etree.SubElement(dst_root, f"{{{NS['ct']}}}Override")
        o.set("PartName", f"/{new_part}")
        o.set("ContentType", ct)


def _new_part_name(dst_root, src_part, counters):
    """Collision-free dst name continuing the family numbering.

    'ppt/slideLayouts/slideLayout3.xml' → 'ppt/slideLayouts/slideLayout12.xml'
    (12 = max existing index in dst + copies already made this run + 1).
    Extensionless filenames keep the whole name as the stem (no mangling).
    """
    d, fname = posixpath.dirname(src_part), posixpath.basename(src_part)
    if "." in fname:
        stem, _, ext = fname.rpartition(".")
    else:
        stem, ext = fname, ""
    prefix = stem.rstrip("0123456789") or "part"
    key = (d, prefix.lower(), ext.lower())
    if key not in counters:
        suffix = (r"\." + re.escape(ext) if ext else "") + r"$"
        rx = re.compile(re.escape(prefix) + r"(\d+)" + suffix, re.IGNORECASE)
        mx = 0
        dst_dir = Path(dst_root) / d
        if dst_dir.is_dir():
            for f in dst_dir.iterdir():
                m = rx.fullmatch(f.name)
                if m:
                    mx = max(mx, int(m.group(1)))
        counters[key] = mx
    counters[key] += 1
    new_fname = (f"{prefix}{counters[key]}.{ext}" if ext
                 else f"{prefix}{counters[key]}")
    return f"{d}/{new_fname}"


def _remove_hlink_elements(xml_path, rids):
    """Strip <a:hlinkClick>/<a:hlinkHover> elements whose r:id is in rids.

    Removes just the hlink element — its parent run/shape is untouched.
    Used after a slide-jump Relationship is dropped, so the slide XML
    doesn't reference an rId missing from its .rels. Returns count removed.
    """
    if not rids:
        return 0
    tree = _xml(xml_path)
    doomed = [
        el
        for tag in ("hlinkClick", "hlinkHover")
        for el in tree.getroot().iter(f"{{{NS['a']}}}{tag}")
        if el.get(f"{{{NS['r']}}}id") in rids
    ]
    for el in doomed:
        el.getparent().remove(el)
    if doomed:
        _write_xml(tree, xml_path)
    return len(doomed)


def _strip_dangling_slide_links(src_dir):
    """Drop slide-jump rels whose Target slide no longer exists.

    Used by extract after slide removals: a kept slide may carry a
    hyperlink jump to a removed slide. Removes the Relationship and the
    matching hlink elements from the slide XML; warns per stripped link.
    """
    root = Path(src_dir)
    slides_dir = root / "ppt" / "slides"
    if not slides_dir.is_dir():
        return
    for slide_xml in slides_dir.glob("slide*.xml"):
        part = f"ppt/slides/{slide_xml.name}"
        rels_p = _part_rels_path(root, part)
        if not rels_p.exists():
            continue
        tree = _xml(rels_p)
        dangling = set()
        for rel in list(_iter_rels(tree)):
            if rel.get("Type") != SLIDE_RT:
                continue
            if rel.get("TargetMode") == "External":
                continue
            target = _resolve_target(part, rel.get("Target"))
            if not (root / target).exists():
                print(f"  WARN: stripping dangling slide link "
                      f"{rel.get('Id')} in {part} (target {target} "
                      f"was removed)", file=sys.stderr)
                dangling.add(rel.get("Id"))
                rel.getparent().remove(rel)
        if dangling:
            _write_xml(tree, rels_p)
            _remove_hlink_elements(slide_xml, dangling)


def _copy_part_tree(state, src_part):
    """Copy src_part and its internal dependency graph into dst.

    Returns the new package path. rIds inside each copied part stay valid
    because its .rels is copied wholesale — only Targets are rewritten.
    Masters are special-cased: their slideLayout relationships are deferred
    and pruned in _finalize_masters to just the layouts copied this run.
    """
    if src_part in state["done"]:
        return state["map"][src_part]
    new_part = state["map"].get(src_part) or _new_part_name(
        state["dst"], src_part, state["counters"])
    state["map"][src_part] = new_part
    state["done"].add(src_part)

    src_file = state["src"] / src_part
    dst_file = state["dst"] / new_part
    dst_file.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(src_file, dst_file)
    _ct_register(state["dst_ct"], state["src_ct"], src_part, new_part)

    src_rels = _part_rels_path(state["src"], src_part)
    if not src_rels.exists():
        return new_part
    tree = _xml(src_rels)
    root = tree.getroot()
    is_master = src_part.startswith("ppt/slideMasters/")
    dropped_rids = set()
    for rel in list(_iter_rels(tree)):
        if rel.get("TargetMode") == "External":
            continue
        rtype = rel.get("Type")
        target = _resolve_target(src_part, rel.get("Target"))
        if rtype == NOTES_RT:
            root.remove(rel)  # notes are dropped on import (documented)
            continue
        if is_master and rtype == LAYOUT_RT:
            state["master_layouts"].setdefault(new_part, []).append(
                (rel, target))
            continue
        if rtype == SLIDE_RT and target not in state["selected"]:
            print(f"  WARN: dropping link from {src_part} to unselected "
                  f"{target}", file=sys.stderr)
            dropped_rids.add(rel.get("Id"))
            root.remove(rel)
            continue
        new_target = _copy_part_tree(state, target)
        rel.set("Target", _relative_target(new_part, new_target))
    if dropped_rids:
        # The copied XML still carries <a:hlinkClick r:id="..."> pointing
        # at the rIds just dropped — strip them so the part stays valid.
        _remove_hlink_elements(dst_file, dropped_rids)
    if is_master:
        state["master_rels"][new_part] = tree  # written in _finalize_masters
    else:
        rels_out = _part_rels_path(state["dst"], new_part)
        rels_out.parent.mkdir(parents=True, exist_ok=True)
        _write_xml(tree, rels_out)
    return new_part


def _max_visual_id(dst_root):
    """Max id across sldMasterIdLst + every master's sldLayoutIdLst.
    Master/layout ids are presentation-unique with floor 2147483648."""
    mx = 2147483647
    pres = _xml(Path(dst_root) / "ppt" / "presentation.xml")
    for el in pres.getroot().findall("p:sldMasterIdLst/p:sldMasterId", NS):
        mx = max(mx, int(el.get("id")))
    for m in (Path(dst_root) / "ppt" / "slideMasters").glob("slideMaster*.xml"):
        for el in _xml(m).getroot().findall("p:sldLayoutIdLst/p:sldLayoutId", NS):
            mx = max(mx, int(el.get("id")))
    return mx


def _finalize_masters(state, idctr):
    """Write deferred master rels; prune+renumber each sldLayoutIdLst.

    A copied master's rels reference ALL its source layouts — keep only the
    layouts actually copied this run, and give surviving entries fresh
    presentation-unique ids. Returns the updated id counter.
    """
    for new_master, tree in state["master_rels"].items():
        kept_rids = set()
        for rel, src_target in state["master_layouts"].get(new_master, []):
            if src_target in state["map"]:
                rel.set("Target", _relative_target(
                    new_master, state["map"][src_target]))
                kept_rids.add(rel.get("Id"))
            else:
                rel.getparent().remove(rel)
        rels_out = _part_rels_path(state["dst"], new_master)
        rels_out.parent.mkdir(parents=True, exist_ok=True)
        _write_xml(tree, rels_out)

        master_p = state["dst"] / new_master
        mt = _xml(master_p)
        lst = mt.getroot().find("p:sldLayoutIdLst", NS)
        if lst is not None:
            for el in list(lst):
                if el.get(f"{{{NS['r']}}}id") in kept_rids:
                    idctr += 1
                    el.set("id", str(idctr))
                else:
                    lst.remove(el)
        _write_xml(mt, master_p)
    return idctr


def append_decks(dst_pptx, src_pptx, slides_spec, output):
    """Merge: copy selected src slides (default all) after dst's last slide.

    Copies each slide with its full dependency graph — layout, master, theme,
    media, charts (+ colors/style/embedded xlsx) — renamed to avoid
    collisions. Source slides keep their own master/theme (faithful import);
    speaker notes are dropped.
    """
    dst_pptx, src_pptx, output = Path(dst_pptx), Path(src_pptx), Path(output)
    with tempfile.TemporaryDirectory() as td:
        dst_root, src_root = Path(td) / "dst", Path(td) / "src"
        unpack(dst_pptx, dst_root)
        unpack(src_pptx, src_root)
        s_pres, s_rels, _ = _presentation_parts(src_root)
        entries = _slide_entries(_xml(s_pres), _xml(s_rels))
        n = len(entries)
        try:
            sel = _parse_selection(slides_spec or f"1-{n}", n)
        except ValueError as e:
            print(f"ERROR: {e}", file=sys.stderr)
            return False

        selected_parts = [f"ppt/slides/{entries[p - 1][2]}" for p in sel]
        state = {
            "src": src_root, "dst": dst_root,
            "map": {}, "done": set(), "counters": {},
            "selected": set(selected_parts),
            "master_rels": {}, "master_layouts": {},
            "src_ct": _ct_index(src_root),
            "dst_ct": _xml(dst_root / "[Content_Types].xml"),
        }
        # Pre-assign slide names so slide→slide links within the selection
        # resolve regardless of copy order.
        for part in selected_parts:
            state["map"][part] = _new_part_name(
                dst_root, part, state["counters"])
        try:
            new_slides = [_copy_part_tree(state, p) for p in selected_parts]
        except (ValueError, OSError, etree.XMLSyntaxError) as e:
            print(f"ERROR: copy failed: {e}", file=sys.stderr)
            return False
        idctr = _finalize_masters(state, _max_visual_id(dst_root))
        _write_xml(state["dst_ct"], dst_root / "[Content_Types].xml")

        # Register new slides + masters in dst presentation.xml(.rels).
        d_pres_p, d_rels_p, _ = _presentation_parts(dst_root)
        pres, rels = _xml(d_pres_p), _xml(d_rels_p)
        rel_root = rels.getroot()
        max_rid = max((int(r.get("Id")[3:]) for r in _iter_rels(rels)
                       if re.fullmatch(r"rId\d+", r.get("Id") or "")),
                      default=0)
        sld_lst = pres.getroot().find("p:sldIdLst", NS)
        max_sldid = max((int(e.get("id")) for e in
                         sld_lst.findall("p:sldId", NS)), default=255)
        for new_slide in new_slides:
            max_rid += 1
            rel = etree.SubElement(rel_root, f"{{{NS['rel']}}}Relationship")
            rel.set("Id", f"rId{max_rid}")
            rel.set("Type", SLIDE_RT)
            rel.set("Target", new_slide[len("ppt/"):])
            max_sldid += 1
            sld = etree.SubElement(sld_lst, f"{{{NS['p']}}}sldId")
            sld.set("id", str(max_sldid))
            sld.set(f"{{{NS['r']}}}id", f"rId{max_rid}")
        master_lst = pres.getroot().find("p:sldMasterIdLst", NS)
        for new_master in state["master_rels"]:
            max_rid += 1
            rel = etree.SubElement(rel_root, f"{{{NS['rel']}}}Relationship")
            rel.set("Id", f"rId{max_rid}")
            rel.set("Type", MASTER_RT)
            rel.set("Target", new_master[len("ppt/"):])
            idctr += 1
            mid = etree.SubElement(master_lst, f"{{{NS['p']}}}sldMasterId")
            mid.set("id", str(idctr))
            mid.set(f"{{{NS['r']}}}id", f"rId{max_rid}")
        _write_xml(rels, d_rels_p)
        _write_xml(pres, d_pres_p)

        if not pack(dst_root, output):
            return False
    print(f"Appended {len(new_slides)} slide(s) from {src_pptx} → {output} "
          f"({len(state['master_rels'])} master(s) imported; notes dropped)")
    return True


A_T = "{http://schemas.openxmlformats.org/drawingml/2006/main}t"


def _slide_xml_files(src_dir):
    """slideN.xml paths sorted by N."""
    slides_dir = Path(src_dir) / "ppt" / "slides"
    files = [p for p in slides_dir.glob("slide*.xml") if p.stem[5:].isdigit()]
    return sorted(files, key=lambda p: int(p.stem[5:]))


def inventory(src_dir):
    """Every text run in every slide as [{slide, run, text}] (document order).

    'run' is the index of the <a:t> within its slide — the stable address
    replace_runs() uses. Returns the list (the CLI prints it as JSON).

    NOTE: 'slide' is the slideN.xml file number, not the deck position —
    after a reorder these differ; run the 'list' subcommand to correlate.
    """
    out = []
    for f in _slide_xml_files(src_dir):
        n = int(f.stem[5:])
        tree = _xml(f)
        for i, t in enumerate(tree.iter(A_T)):
            out.append({"slide": n, "run": i, "text": t.text or ""})
    return out


def replace_runs(src_dir, edits_json):
    """Apply [{slide, run, text}] edits; formatting (rPr) is untouched.

    NOTE: 'slide' is the slideN.xml file number, not the deck position —
    after a reorder these differ; run the 'list' subcommand to correlate.

    On error, slides processed before the failure are already written;
    re-unpack from the original .pptx to reset.
    """
    import json
    edits = json.loads(Path(edits_json).read_text(encoding="utf-8"))
    by_slide = {}
    for e in edits:
        by_slide.setdefault(int(e["slide"]), {})[int(e["run"])] = e["text"]
    for f in _slide_xml_files(src_dir):
        n = int(f.stem[5:])
        if n not in by_slide:
            continue
        tree = _xml(f)
        runs = list(tree.iter(A_T))
        for idx, new_text in by_slide[n].items():
            if idx >= len(runs):
                raise SystemExit(
                    f"slide {n}: run {idx} out of range (has {len(runs)})")
            runs[idx].text = new_text
        _write_xml(tree, f)
        print(f"slide {n}: {len(by_slide[n])} run(s) replaced")


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    sub = parser.add_subparsers(dest="cmd", required=True)
    p_un = sub.add_parser("unpack"); p_un.add_argument("pptx"); p_un.add_argument("dir")
    p_pk = sub.add_parser("pack"); p_pk.add_argument("dir"); p_pk.add_argument("pptx")
    p_ls = sub.add_parser("list"); p_ls.add_argument("dir")
    p_du = sub.add_parser("duplicate"); p_du.add_argument("dir"); p_du.add_argument(
        "n", type=int, help="1-based deck position (matches list output)")
    p_rm = sub.add_parser("remove"); p_rm.add_argument("dir"); p_rm.add_argument(
        "n", type=int, help="1-based deck position (matches list output)")
    p_ro = sub.add_parser("reorder"); p_ro.add_argument("dir"); p_ro.add_argument(
        "order", help="Comma-separated positions, e.g. 3,1,2,4")
    p_cl = sub.add_parser("clean"); p_cl.add_argument("dir")
    p_inv = sub.add_parser("inventory", help="JSON dump of every text run (addresses are slideN.xml file numbers)"); p_inv.add_argument("dir")
    p_rep = sub.add_parser("replace"); p_rep.add_argument("dir")
    p_rep.add_argument("edits", help="JSON: [{slide, run, text}, ...]")
    p_ex = sub.add_parser("extract", help="Split: new deck from selected slides")
    p_ex.add_argument("pptx")
    p_ex.add_argument("selection", help='1-based positions: "3-7", "2,5,9", "1,3-5"')
    p_ex.add_argument("--output", required=True)
    p_ap = sub.add_parser("append", help="Merge: copy src slides after dst's last slide")
    p_ap.add_argument("dst")
    p_ap.add_argument("src")
    p_ap.add_argument("--slides", default=None,
                      help='Positions in src to copy (default: all), e.g. "2-4"')
    p_ap.add_argument("--output", required=True)
    args = parser.parse_args()

    ok = True
    if args.cmd == "unpack":
        unpack(args.pptx, args.dir)
    elif args.cmd == "pack":
        ok = pack(args.dir, args.pptx)
    elif args.cmd == "list":
        list_slides(args.dir)
    elif args.cmd == "duplicate":
        ok = duplicate(args.dir, args.n)
    elif args.cmd == "remove":
        ok = remove(args.dir, args.n)
    elif args.cmd == "reorder":
        ok = reorder(args.dir, args.order)
    elif args.cmd == "clean":
        ok = clean_orphans(args.dir)
    elif args.cmd == "inventory":
        import json
        print(json.dumps(inventory(args.dir), indent=2, ensure_ascii=False))
    elif args.cmd == "replace":
        replace_runs(args.dir, args.edits)
    elif args.cmd == "extract":
        ok = extract(args.pptx, args.selection, args.output)
    elif args.cmd == "append":
        ok = append_decks(args.dst, args.src, args.slides, args.output)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
