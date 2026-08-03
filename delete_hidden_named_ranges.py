#!/usr/bin/env python3
"""Delete hidden named ranges from an Excel workbook without rebuilding the file.

Hidden names are the ones VBA creates with ``Name.Visible = False``; in the file
they are ``<definedName hidden="1">`` entries in ``xl/workbook.xml``. Add-ins,
crashed merges and years of copied sheets leave thousands of them behind, and
Excel's Name Manager will not even show them.

The usual advice -- load the book in openpyxl, delete names, save -- rewrites
every part of the package, which silently drops charts, pivot tables, slicers,
images, form controls and the VBA project. This script instead opens the .xlsx
as the zip archive it is, edits the bytes of the one XML part that holds the
names, and copies every other entry through verbatim (same order, same
compression, same timestamps). Nothing else in the workbook can change.

Safety rails, all on by default:
  * Excel's own built-ins (``_xlnm.Print_Area``, ``_xlnm._FilterDatabase``, ...)
    are hidden by design and load-bearing, so they are kept.
  * A name still referenced by a formula, chart, table, data validation rule or
    the VBA project is kept -- deleting it turns those into #NAME? errors.
  * Output goes to a new file; --in-place keeps a .bak alongside.

Usage:
    python delete_hidden_named_ranges.py book.xlsm --dry-run
    python delete_hidden_named_ranges.py book.xlsm -o cleaned.xlsm
    python delete_hidden_named_ranges.py book.xlsm --in-place
    python delete_hidden_named_ranges.py book.xlsm --keep 'rng_*' --keep MyName

Works on .xlsx/.xlsm/.xltx/.xltm. No third-party packages, Python 3.8+.
"""

from __future__ import annotations

import argparse
import fnmatch
import os
import re
import shutil
import sys
import tempfile
import zipfile
from typing import Dict, List, NamedTuple, Optional, Sequence, Tuple
from xml.sax.saxutils import unescape

WORKBOOK_PART = "xl/workbook.xml"

# Excel's reserved names. They live under the _xlnm namespace, several are
# always flagged hidden, and removing them breaks print ranges and autofilters.
BUILTIN_PREFIX = "_xlnm."

# Attribute values Excel (or another writer) may use for a true boolean.
TRUE_VALUES = {"1", "true", "True", "TRUE"}

DEFINED_NAMES_BLOCK = re.compile(rb"<definedNames(?:\s[^>]*)?>.*?</definedNames\s*>", re.S)
EMPTY_DEFINED_NAMES = re.compile(rb"<definedNames(?:\s[^>]*)?/>")
# A single <definedName ...>refersTo</definedName>, plus any whitespace ahead of
# it so removing an entry does not leave a ragged blank line behind.
DEFINED_NAME_ENTRY = re.compile(
    rb"\s*<definedName(?:\s[^>]*)?(?:/>|>.*?</definedName\s*>)", re.S
)
ATTRIBUTE = re.compile(rb"""([A-Za-z_:][-\w:.]*)\s*=\s*"([^"]*)"|([A-Za-z_:][-\w:.]*)\s*=\s*'([^']*)'""")
ELEMENT_TEXT = re.compile(rb"^<[^>]*>(.*)</[^>]+>$", re.S)


class DefinedName(NamedTuple):
    """One <definedName> entry, kept together with its exact source bytes."""

    name: str
    refers_to: str
    hidden: bool
    local_sheet_id: Optional[str]
    raw: bytes
    span: Tuple[int, int]

    @property
    def is_builtin(self) -> bool:
        return self.name.startswith(BUILTIN_PREFIX)

    @property
    def scope(self) -> str:
        return "workbook" if self.local_sheet_id is None else f"sheet #{self.local_sheet_id}"


class Decision(NamedTuple):
    entry: DefinedName
    delete: bool
    reason: str


# --------------------------------------------------------------------------- #
# Parsing
# --------------------------------------------------------------------------- #

def parse_defined_names(workbook_xml: bytes) -> Tuple[Optional[Tuple[int, int]], List[DefinedName]]:
    """Locate the <definedNames> block and every entry inside it.

    Offsets are absolute into ``workbook_xml`` so edits can be applied by slicing
    the original bytes; the XML is never re-serialised.
    """
    block = DEFINED_NAMES_BLOCK.search(workbook_xml)
    if block is None:
        return None, []

    block_start, block_end = block.span()
    entries: List[DefinedName] = []
    for match in DEFINED_NAME_ENTRY.finditer(workbook_xml, block_start, block_end):
        raw = match.group(0)
        tag_end = raw.index(b">") + 1
        attrs = _parse_attributes(raw[:tag_end])

        name = attrs.get("name", "")
        if not name:
            # Nameless entry: malformed, leave it strictly alone.
            continue

        text_match = ELEMENT_TEXT.match(raw.strip())
        refers_to = _decode(text_match.group(1)) if text_match else ""

        entries.append(
            DefinedName(
                name=name,
                refers_to=refers_to.strip(),
                hidden=attrs.get("hidden", "") in TRUE_VALUES,
                local_sheet_id=attrs.get("localSheetId"),
                raw=raw,
                span=match.span(),
            )
        )
    return (block_start, block_end), entries


def _parse_attributes(tag: bytes) -> Dict[str, str]:
    attrs: Dict[str, str] = {}
    for m in ATTRIBUTE.finditer(tag):
        key = m.group(1) or m.group(3)
        value = m.group(2) if m.group(1) else m.group(4)
        attrs[_decode(key)] = _decode(value)
    return attrs


def _decode(raw: bytes) -> str:
    text = raw.decode("utf-8", errors="replace")
    return unescape(text, {"&quot;": '"', "&apos;": "'"})


# --------------------------------------------------------------------------- #
# Usage detection
# --------------------------------------------------------------------------- #

def find_used_names(
    archive: zipfile.ZipFile,
    candidates: Sequence[DefinedName],
    workbook_xml: bytes,
    names_block: Optional[Tuple[int, int]],
) -> Dict[str, str]:
    """Return {name: where it is used} for candidates referenced elsewhere.

    Scans every XML part (sheets, charts, tables, pivot caches, data validation)
    and the compiled VBA project. A defined name is a bare token in a formula, so
    matching is token-based: bounded by characters that cannot appear mid-name.
    This errs towards false positives -- keeping a name costs nothing, deleting a
    live one costs a #NAME? error.
    """
    if not candidates:
        return {}

    patterns = {
        entry.name: re.compile(
            r"(?<![A-Za-z0-9_.\\])" + re.escape(entry.name) + r"(?![A-Za-z0-9_.])"
        )
        for entry in candidates
    }
    used: Dict[str, str] = {}

    for info in archive.infolist():
        if info.is_dir() or len(used) == len(patterns):
            continue

        part = info.filename
        if part.endswith(".xml") or part.endswith(".rels"):
            data = archive.read(part)
            if part == WORKBOOK_PART and names_block is not None:
                # Skip the definitions themselves; a name always matches its own
                # entry. Cross-references from *other* names still count.
                data = data[: names_block[0]] + data[names_block[1] :]
            haystacks = [(data.decode("utf-8", errors="replace"), part)]
        elif part.endswith("vbaProject.bin"):
            blob = archive.read(part)
            # VBA stores identifiers as ASCII in compressed module streams and as
            # UTF-16LE in name tables; check both, cheaply.
            haystacks = [
                (blob.decode("latin-1", errors="replace"), part),
                (blob.decode("utf-16-le", errors="replace"), part + " (UTF-16)"),
            ]
        else:
            continue

        for text, label in haystacks:
            for name, pattern in patterns.items():
                if name not in used and pattern.search(text):
                    used[name] = label

    return used


# --------------------------------------------------------------------------- #
# Policy
# --------------------------------------------------------------------------- #

def decide(
    entries: Sequence[DefinedName],
    used: Dict[str, str],
    keep_patterns: Sequence[str],
    only_patterns: Sequence[str],
    include_builtin: bool,
    force: bool,
) -> List[Decision]:
    decisions: List[Decision] = []
    for entry in entries:
        if not entry.hidden:
            decisions.append(Decision(entry, False, "visible"))
        elif entry.is_builtin and not include_builtin:
            decisions.append(Decision(entry, False, "Excel built-in (--include-builtin to remove)"))
        elif any(fnmatch.fnmatch(entry.name, p) for p in keep_patterns):
            decisions.append(Decision(entry, False, "matched --keep"))
        elif only_patterns and not any(fnmatch.fnmatch(entry.name, p) for p in only_patterns):
            decisions.append(Decision(entry, False, "not matched by --only"))
        elif entry.name in used and not force:
            decisions.append(Decision(entry, False, f"referenced by {used[entry.name]} (--force to remove)"))
        else:
            decisions.append(Decision(entry, True, "hidden"))
    return decisions


def apply_deletions(workbook_xml: bytes, block: Tuple[int, int], decisions: Sequence[Decision]) -> bytes:
    """Cut the doomed entries out of the raw XML, back to front.

    Editing from the end keeps every earlier offset valid, so no byte outside the
    removed ranges is disturbed.
    """
    doomed = [d.entry for d in decisions if d.delete]
    if not doomed:
        return workbook_xml

    result = workbook_xml
    for entry in sorted(doomed, key=lambda e: e.span[0], reverse=True):
        start, end = entry.span
        result = result[:start] + result[end:]

    # An empty <definedNames></definedNames> violates the schema (Excel wants at
    # least one child and will report the file as unreadable), so drop the
    # container itself once the last name inside it is gone.
    leftover = DEFINED_NAMES_BLOCK.search(result)
    if leftover and not DEFINED_NAME_ENTRY.search(leftover.group(0)):
        result = result[: leftover.start()] + result[leftover.end() :]
    return EMPTY_DEFINED_NAMES.sub(b"", result)


# --------------------------------------------------------------------------- #
# Repackaging
# --------------------------------------------------------------------------- #

def rewrite_archive(source: str, destination: str, replacements: Dict[str, bytes]) -> None:
    """Copy the package entry for entry, swapping in the edited parts.

    Order, per-entry compression, timestamps, external attributes and the archive
    comment are all preserved, so the result differs from the input only in the
    parts we rewrote.
    """
    with zipfile.ZipFile(source, "r") as src, zipfile.ZipFile(destination, "w") as dst:
        dst.comment = src.comment
        for info in src.infolist():
            data = replacements.get(info.filename)
            if data is None:
                data = src.read(info.filename)

            # Carry the original metadata over rather than letting zipfile
            # synthesise a fresh header.
            out = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            out.compress_type = info.compress_type
            out.external_attr = info.external_attr
            out.internal_attr = info.internal_attr
            out.create_system = info.create_system
            out.comment = info.comment
            dst.writestr(out, data)


def verify(path: str, expected_removed: Sequence[str]) -> List[str]:
    """Re-open the written file and confirm it is intact. Returns problems found."""
    problems: List[str] = []
    try:
        with zipfile.ZipFile(path, "r") as archive:
            bad = archive.testzip()
            if bad is not None:
                problems.append(f"corrupt entry in output: {bad}")
            names = set(archive.namelist())
            for required in ("[Content_Types].xml", WORKBOOK_PART):
                if required not in names:
                    problems.append(f"missing required part: {required}")

            # Every XML part must still parse.
            from xml.etree import ElementTree

            for part in sorted(names):
                if part.endswith(".xml") or part.endswith(".rels"):
                    try:
                        ElementTree.fromstring(archive.read(part))
                    except ElementTree.ParseError as exc:
                        problems.append(f"{part} is not well-formed XML: {exc}")

            _, remaining = parse_defined_names(archive.read(WORKBOOK_PART))
            still_there = {e.name for e in remaining} & set(expected_removed)
            if still_there:
                problems.append("names survived deletion: " + ", ".join(sorted(still_there)))
    except zipfile.BadZipFile as exc:
        problems.append(f"output is not a readable zip: {exc}")
    return problems


# --------------------------------------------------------------------------- #
# Reporting / CLI
# --------------------------------------------------------------------------- #

def report(decisions: Sequence[Decision], verbose: bool) -> None:
    hidden = [d for d in decisions if d.entry.hidden]
    doomed = [d for d in hidden if d.delete]
    kept = [d for d in hidden if not d.delete]

    print(f"{len(decisions)} defined name(s): {len(hidden)} hidden, {len(doomed)} to delete")

    if doomed:
        print("\nDeleting:")
        for d in doomed:
            print(f"  - {d.entry.name}  [{d.entry.scope}]  ->  {_truncate(d.entry.refers_to)}")

    if kept:
        print("\nKeeping (hidden but protected):")
        for d in kept:
            print(f"  - {d.entry.name}  [{d.entry.scope}]  ({d.reason})")

    if verbose:
        visible = [d for d in decisions if not d.entry.hidden]
        if visible:
            print("\nVisible names (untouched):")
            for d in visible:
                print(f"  - {d.entry.name}  [{d.entry.scope}]  ->  {_truncate(d.entry.refers_to)}")


def _truncate(text: str, limit: int = 70) -> str:
    text = " ".join(text.split())
    return text if len(text) <= limit else text[: limit - 1] + "…"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Delete hidden named ranges from an Excel workbook, in place in the zip package.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="The workbook is edited at the XML level; charts, pivot tables, images\n"
        "and the VBA project are copied through byte for byte.",
    )
    parser.add_argument("workbook", help="Path to the .xlsx/.xlsm/.xltx/.xltm file")
    parser.add_argument("-o", "--out", help="Write result here (default: <name>.cleaned<ext>)")
    parser.add_argument("--in-place", action="store_true", help="Overwrite the input, keeping a .bak copy")
    parser.add_argument("-n", "--dry-run", action="store_true", help="Report what would be deleted, write nothing")
    parser.add_argument("--keep", action="append", default=[], metavar="PATTERN",
                        help="Never delete names matching this glob (repeatable)")
    parser.add_argument("--only", action="append", default=[], metavar="PATTERN",
                        help="Only delete names matching this glob (repeatable)")
    parser.add_argument("--include-builtin", action="store_true",
                        help="Also delete _xlnm.* names (print areas, filter ranges) -- rarely what you want")
    parser.add_argument("--force", action="store_true",
                        help="Delete hidden names even when something still references them")
    parser.add_argument("--no-backup", action="store_true", help="Skip the .bak copy with --in-place")
    parser.add_argument("-v", "--verbose", action="store_true", help="Also list visible names")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)

    source = args.workbook
    if not os.path.isfile(source):
        print(f"error: no such file: {source}", file=sys.stderr)
        return 2
    if args.out and args.in_place:
        print("error: --out and --in-place are mutually exclusive", file=sys.stderr)
        return 2

    try:
        with zipfile.ZipFile(source, "r") as archive:
            try:
                workbook_xml = archive.read(WORKBOOK_PART)
            except KeyError:
                print(f"error: {source} has no {WORKBOOK_PART} -- is it really a modern Excel file?\n"
                      "       Legacy .xls is a different format; re-save it as .xlsx/.xlsm first.",
                      file=sys.stderr)
                return 2

            block, entries = parse_defined_names(workbook_xml)
            hidden = [e for e in entries if e.hidden]
            candidates = [
                e for e in hidden
                if (args.include_builtin or not e.is_builtin)
                and not any(fnmatch.fnmatch(e.name, p) for p in args.keep)
            ]
            used = {} if args.force else find_used_names(archive, candidates, workbook_xml, block)
    except zipfile.BadZipFile:
        print(f"error: {source} is not a valid Excel package (zip). If it is a legacy .xls,\n"
              "       re-save it as .xlsx/.xlsm first.", file=sys.stderr)
        return 2

    decisions = decide(entries, used, args.keep, args.only, args.include_builtin, args.force)
    report(decisions, args.verbose)

    doomed = [d.entry.name for d in decisions if d.delete]
    if not doomed:
        print("\nNothing to delete; file left untouched.")
        return 0
    if args.dry_run:
        print("\nDry run: no file written.")
        return 0

    assert block is not None  # names exist, so the block does too
    new_workbook_xml = apply_deletions(workbook_xml, block, decisions)

    if args.in_place:
        destination = source
        if not args.no_backup:
            backup = source + ".bak"
            shutil.copy2(source, backup)
            print(f"\nBackup: {backup}")
    elif args.out:
        destination = args.out
    else:
        stem, ext = os.path.splitext(source)
        destination = f"{stem}.cleaned{ext}"

    # Build the new package beside the destination, verify it, and only then move
    # it into place -- an interrupted run can never leave a half-written workbook.
    target_dir = os.path.dirname(os.path.abspath(destination)) or "."
    handle, staging = tempfile.mkstemp(dir=target_dir, suffix=".tmp")
    os.close(handle)
    try:
        rewrite_archive(source, staging, {WORKBOOK_PART: new_workbook_xml})
        problems = verify(staging, doomed)
        if problems:
            print("\nAborted -- output failed verification, original untouched:", file=sys.stderr)
            for problem in problems:
                print(f"  * {problem}", file=sys.stderr)
            return 1
        os.replace(staging, destination)
    finally:
        if os.path.exists(staging):
            os.remove(staging)

    print(f"\nRemoved {len(doomed)} hidden name(s) -> {destination}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
