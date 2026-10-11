"""Embed the Excel template into the script package as template_data.py.

Run this again whenever the template changes (new yearly limits, new year sheet...):
    python embed_template.py                          # uses template_source_for_embedding from config.json
    python embed_template.py "my_template.xltx"
    python embed_template.py --share                  # also remove stored paths, labels, printer settings, tracking ids
    python embed_template.py --strip-author           # blank the author / last-modified-by names first
"""
import argparse
import base64
import io
import json
import re
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent


def strip_author(data: bytes) -> bytes:
    src = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            content = src.read(item.filename)
            if item.filename == "docProps/core.xml":
                text = content.decode("utf-8")
                text = re.sub(r"(<dc:creator>).*?(</dc:creator>)", r"\1\2", text)
                text = re.sub(r"(<cp:lastModifiedBy>).*?(</cp:lastModifiedBy>)", r"\1\2", text)
                content = text.encode("utf-8")
            dst.writestr(item, content)
    return out.getvalue()


def _strip_tracking_ids(content: bytes) -> bytes:
    """Drop revision/tracking attributes and elements (random per-document ids; no effect on the content)."""
    text = content.decode("utf-8")
    text = re.sub(r"<xr:revisionPtr\b[^>]*/>", "", text)
    text = re.sub(r'\sxr\d*:(?:uid|uidLastSave|coauthVersionLast|coauthVersionMax|revIDLastSave)="[^"]*"', "", text)
    return text.encode("utf-8")


def clean_for_sharing(data: bytes) -> bytes:
    """Remove everything in the template that can identify its origin:
      - the stored file path (user name + folder)                      xl/workbook.xml
      - the sensitivity-label properties (organization ids)           docProps/custom.xml
      - printer settings (print-server / queue names, user names)      xl/printerSettings/*.bin
      - Office add-in panes (add-in ids, a per-file id of the add-in)  xl/webextensions/*
      - per-document revision / tracking ids (xr:uid, documentId ...) in every XML part
    The zip container is rewritten with fixed timestamps, no comments and no extra fields.
    The author name in docProps/core.xml is left alone (use --strip-author to blank it)."""
    src = zipfile.ZipFile(io.BytesIO(data))
    printer_parts = {n for n in src.namelist() if n.startswith("xl/printerSettings/")}
    # relationship ids that point at printer settings, per worksheet rels file
    printer_rids = {}
    for n in src.namelist():
        if n.startswith("xl/worksheets/_rels/") and n.endswith(".rels"):
            text = src.read(n).decode("utf-8")
            ids = re.findall(r'<Relationship [^>]*Id="([^"]+)"[^>]*Target="[^"]*printerSettings[^"]*"[^>]*/>', text)
            ids += re.findall(r'<Relationship [^>]*Target="[^"]*printerSettings[^"]*"[^>]*Id="([^"]+)"[^>]*/>', text)
            sheet = "xl/worksheets/" + n.rsplit("/", 1)[1][: -len(".rels")]
            printer_rids[sheet] = set(ids)

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            name = item.filename
            if name == "docProps/custom.xml" or name in printer_parts or name.startswith("xl/webextensions/"):
                continue
            content = src.read(name)
            if name.endswith((".xml", ".rels")):
                content = _strip_tracking_ids(content)
            if name == "[Content_Types].xml":
                text = content.decode("utf-8")
                text = re.sub(r'<Override PartName="/docProps/custom\.xml"[^>]*/>', "", text)
                text = re.sub(r'<Override PartName="/xl/printerSettings/[^"]*"[^>]*/>', "", text)
                text = re.sub(r'<Override PartName="/xl/webextensions/[^"]*"[^>]*/>', "", text)
                if not any(n.endswith(".bin") and n not in printer_parts for n in src.namelist()):
                    text = re.sub(r'<Default Extension="bin"[^>]*/>', "", text)
                content = text.encode("utf-8")
            elif name == "_rels/.rels":
                text = content.decode("utf-8")
                text = re.sub(r'<Relationship [^>]*custom-properties[^>]*/>', "", text)
                text = re.sub(r'<Relationship [^>]*webextension[^>]*/>', "", text)
                content = text.encode("utf-8")
            elif name.startswith("xl/worksheets/_rels/") and name.endswith(".rels"):
                text = content.decode("utf-8")
                text = re.sub(r'<Relationship [^>]*printerSettings[^>]*/>', "", text)
                content = text.encode("utf-8")
            elif name == "xl/workbook.xml":
                text = content.decode("utf-8")
                text = re.sub(r'<mc:AlternateContent[^>]*><mc:Choice Requires="x15"><x15ac:absPath [^>]*/></mc:Choice></mc:AlternateContent>', "", text)
                content = text.encode("utf-8")
            elif name in printer_rids and printer_rids[name]:
                text = content.decode("utf-8")
                for rid in printer_rids[name]:
                    text = text.replace(' r:id="%s"' % rid, "")      # <pageSetup ... r:id=".."/>
                content = text.encode("utf-8")
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 0
            dst.writestr(info, content)
    return out.getvalue()


def clear_input_values(data: bytes, cfg: dict):
    """Blank leftover test values in the cells the program (or the user) fills (the rows of config.json) of every
    year sheet. Formulas are never touched. Returns (new_data, ["2026!G23", ...])."""
    from openpyxl.utils import column_index_from_string, get_column_letter

    wcfg = cfg["workbook"]
    rows = sorted(set(wcfg["rows"].values()))
    first = column_index_from_string(wcfg["first_month_column"])
    cols = [get_column_letter(first + i) for i in range(12)]

    src = zipfile.ZipFile(io.BytesIO(data))
    workbook = src.read("xl/workbook.xml").decode("utf-8")
    rels = src.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    targets = {}
    for rel in re.findall(r"<Relationship\b[^>]*>", rels):
        rid, target = re.search(r'\bId="([^"]+)"', rel), re.search(r'\bTarget="([^"]+)"', rel)
        if rid and target:
            targets[rid.group(1)] = target.group(1)
    sheet_files = {}
    for tag in re.findall(r"<sheet\b[^>]*>", workbook):
        name, rid = re.search(r'\bname="([^"]*)"', tag), re.search(r'\br:id="([^"]+)"', tag)
        if name and rid and rid.group(1) in targets:
            sheet_files["xl/" + targets[rid.group(1)].lstrip("/").replace("xl/", "", 1)] = name.group(1)

    cleared = []
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            content = src.read(item.filename)
            sheet = sheet_files.get(item.filename)
            if sheet in wcfg["year_sheets"]:
                text = content.decode("utf-8")
                for row in rows:
                    for col in cols:
                        ref = "%s%d" % (col, row)
                        m = re.search(r'<c r="%s"((?:\s[^>]*?)?)(?<!/)>(.*?)</c>' % ref, text, re.S)
                        if not m or "<f" in m.group(2) or 't="s"' in m.group(1) or "<v>" not in m.group(2):
                            continue                          # empty, a formula, or text: leave it alone
                        attrs = re.sub(r'\st="[^"]*"', "", m.group(1))
                        text = text[:m.start()] + '<c r="%s"%s/>' % (ref, attrs) + text[m.end():]
                        cleared.append("%s!%s" % (sheet, ref))
                content = text.encode("utf-8")
            info = zipfile.ZipInfo(item.filename, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 0
            dst.writestr(info, content)
    return out.getvalue(), cleared


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("template", nargs="?", help="path to the .xltx/.xlsx template (default: template_source_for_embedding from config.json)")
    ap.add_argument("--strip-author", action="store_true", help="blank the author names in the file properties")
    ap.add_argument("--write-xltx", help="also save the (cleaned) template itself to this path")
    ap.add_argument("--out-dir", help="write template_data.py here instead of next to this script")
    ap.add_argument("--share", action="store_true",
                    help="remove stored file path, sensitivity label and printer settings (author name is kept)")
    args = ap.parse_args()

    cfg = json.loads((HERE / "config.json").read_text(encoding="utf-8"))
    path = Path(args.template or cfg.get("template_source_for_embedding") or "")
    if not path.is_file():
        print("ERROR: template not found: %s" % path)
        return 2
    data = path.read_bytes()
    cfg_for_clear = json.loads((HERE / "config.json").read_text(encoding="utf-8"))
    data, cleared = clear_input_values(data, cfg_for_clear)
    print("blanked leftover input values: %s" % (", ".join(cleared) or "none"))
    if args.share:
        data = clean_for_sharing(data)
    if args.strip_author:
        data = strip_author(data)

    if args.write_xltx:
        Path(args.write_xltx).write_bytes(data)
        print("saved cleaned template -> %s (%d bytes)" % (args.write_xltx, len(data)))
    b64 = base64.b64encode(data).decode("ascii")
    lines = [b64[i:i + 100] for i in range(0, len(b64), 100)]
    body = "\n".join('    "%s"' % l for l in lines)
    (Path(args.out_dir) / "template_data.py" if args.out_dir else HERE / "template_data.py").write_text(
        '"""Generated by embed_template.py - do not edit by hand.\nSource: %s (%d bytes)"""\nTEMPLATE_B64 = (\n%s\n)\n'
        % (path.name, len(data), body), encoding="utf-8")
    print("embedded %s (%d bytes, share-clean: %s, author stripped: %s) -> template_data.py"
          % (path.name, len(data), args.share, args.strip_author))
    return 0


if __name__ == "__main__":
    sys.exit(main())
