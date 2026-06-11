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

Workflow: unpack → (structural ops: duplicate/remove) → edit slide XML text →
pack → QA with qa_check.py + render_slides.py. See references/editing.md.
"""
import argparse
import re
import shutil
import sys
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
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
