# /// script
# requires-python = ">=3.12"
# dependencies = [
#     "openpyxl",
# ]
# ///
"""
Roundtrip fidelity demo: openpyxl vs offidized.

Opens an xlsx file with a bar chart, changes ONE cell, saves it back.
Then compares what each library preserved vs destroyed.

Usage:
    uv run demo/roundtrip_fidelity.py

Output files land in demo/output/ so you can open them in Excel.
"""

import shutil
import subprocess
import zipfile
from pathlib import Path

SOURCE = Path(__file__).resolve().parent / "chart_demo.xlsx"

CYAN = "\033[96m"
GREEN = "\033[92m"
RED = "\033[91m"
YELLOW = "\033[93m"
BOLD = "\033[1m"
DIM = "\033[2m"
RESET = "\033[0m"


def part_inventory(path: Path) -> dict[str, int]:
    with zipfile.ZipFile(path) as zf:
        return {info.filename: info.file_size for info in zf.infolist()}


def roundtrip_openpyxl(src: Path, dst: Path) -> None:
    import openpyxl

    wb = openpyxl.load_workbook(src)
    ws = wb.active
    ws["A1"] = "modified by openpyxl"
    wb.save(dst)


def roundtrip_offidized(src: Path, dst: Path) -> None:
    shutil.copy2(src, dst)
    result = subprocess.run(
        [
            "cargo",
            "run",
            "-p",
            "offidized-cli",
            "--quiet",
            "--",
            "set",
            str(dst),
            "Sales!A1",
            "modified by offidized",
            "-i",
        ],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip())


def print_header(text: str) -> None:
    print(f"\n{BOLD}{CYAN}{'=' * 60}")
    print(f"  {text}")
    print(f"{'=' * 60}{RESET}\n")


def compare(label: str, original: Path, modified: Path) -> None:
    orig_parts = part_inventory(original)
    mod_parts = part_inventory(modified)

    orig_set = set(orig_parts)
    mod_set = set(mod_parts)

    lost = orig_set - mod_set
    added = mod_set - orig_set

    print(f"  {BOLD}{label}{RESET}")
    print(f"  {'─' * 50}")
    print(f"  Original parts:  {len(orig_set)}")
    print(f"  Output parts:    {len(mod_set)}")

    if lost:
        print(f"\n  {RED}{BOLD}LOST ({len(lost)} parts):{RESET}")
        for p in sorted(lost):
            print(f"    {RED}✗ {p}{RESET}")
    else:
        print(f"\n  {GREEN}No parts lost!{RESET}")

    if added:
        print(f"\n  {YELLOW}Added ({len(added)} parts):{RESET}")
        for p in sorted(added):
            print(f"    {YELLOW}+ {p}{RESET}")

    # Check chart parts
    chart_parts = [p for p in orig_set if "chart" in p.lower()]
    if chart_parts:
        missing = [p for p in chart_parts if p not in mod_set]
        if not missing:
            print(
                f"\n  {GREEN}{BOLD}Charts preserved!{RESET} {DIM}({', '.join(sorted(chart_parts))}){RESET}"
            )
        else:
            print(f"\n  {RED}{BOLD}Charts DESTROYED!{RESET}")
            for p in sorted(missing):
                print(f"    {RED}✗ {p}{RESET}")

    # Check drawing parts
    drawing_parts = [p for p in orig_set if "drawing" in p.lower()]
    if drawing_parts:
        missing = [p for p in drawing_parts if p not in mod_set]
        if not missing:
            print(
                f"  {GREEN}Drawings preserved!{RESET} {DIM}({', '.join(sorted(drawing_parts))}){RESET}"
            )
        else:
            print(f"  {RED}{BOLD}Drawings DESTROYED!{RESET}")
            for p in sorted(missing):
                print(f"    {RED}✗ {p}{RESET}")

    # Check if modified sheet still has <drawing> reference
    sheet_xml = None
    try:
        with zipfile.ZipFile(modified) as zf:
            sheet_xml = zf.read("xl/worksheets/sheet1.xml").decode()
    except (KeyError, UnicodeDecodeError):
        pass
    if sheet_xml is not None:
        if "drawing" in sheet_xml:
            print(f"  {GREEN}Sheet→drawing link intact{RESET}")
        else:
            print(f"  {RED}{BOLD}Sheet→drawing link BROKEN{RESET} (chart won't render)")

    print()


def main() -> None:
    if not SOURCE.exists():
        print(f"{RED}Source file not found: {SOURCE}{RESET}")
        print("Run this first to create it:")
        print("  uv run --with openpyxl python3 -c 'import openpyxl; ...'")
        return

    print_header("Roundtrip Fidelity: openpyxl vs offidized")
    print(f"  Source: {DIM}{SOURCE.name}{RESET}")
    print(f"  Parts:  {len(part_inventory(SOURCE))}")
    print("  Task:   Open file with bar chart, change cell A1, save")

    out_dir = Path(__file__).resolve().parent / "output"
    out_dir.mkdir(exist_ok=True)
    original_copy = out_dir / "original.xlsx"
    openpyxl_out = out_dir / "openpyxl_output.xlsx"
    offidized_out = out_dir / "offidized_output.xlsx"

    shutil.copy2(SOURCE, original_copy)

    print_header("Running openpyxl roundtrip...")
    try:
        roundtrip_openpyxl(SOURCE, openpyxl_out)
        print(f"  {GREEN}Saved to {openpyxl_out.name}{RESET}")
    except Exception as e:
        print(f"  {RED}Failed: {e}{RESET}")
        return

    print_header("Running offidized roundtrip...")
    try:
        roundtrip_offidized(SOURCE, offidized_out)
        print(f"  {GREEN}Saved to {offidized_out.name}{RESET}")
    except Exception as e:
        print(f"  {RED}Failed: {e}{RESET}")
        return

    print_header("Results")
    compare("openpyxl", SOURCE, openpyxl_out)
    compare("offidized", SOURCE, offidized_out)

    orig_count = len(part_inventory(SOURCE))
    openpyxl_lost = len(set(part_inventory(SOURCE)) - set(part_inventory(openpyxl_out)))
    offidized_lost = len(
        set(part_inventory(SOURCE)) - set(part_inventory(offidized_out))
    )

    print_header("Verdict")
    print(
        f"  {BOLD}openpyxl:{RESET}  {RED if openpyxl_lost else GREEN}{openpyxl_lost} parts lost{RESET} out of {orig_count}"
    )
    print(
        f"  {BOLD}offidized:{RESET} {RED if offidized_lost else GREEN}{offidized_lost} parts lost{RESET} out of {orig_count}"
    )

    if openpyxl_lost > offidized_lost:
        print(
            f"\n  {GREEN}{BOLD}offidized wins.{RESET} {DIM}Your charts survive.{RESET}\n"
        )
    elif openpyxl_lost == offidized_lost == 0:
        if True:
            # Check deeper: did the sheet→drawing link survive?
            with zipfile.ZipFile(offidized_out) as zf:
                sheet = zf.read("xl/worksheets/sheet1.xml").decode()
            if "drawing" not in sheet:
                print(
                    f"\n  {YELLOW}{BOLD}offidized preserved all parts but broke the sheet→drawing link.{RESET}"
                )
                print(
                    f"  {DIM}The chart XML is there but Excel can't find it.{RESET}\n"
                )
            else:
                print(f"\n  {GREEN}{BOLD}Both preserved everything.{RESET}\n")
    else:
        print()

    print(f"  {DIM}Open these in Excel to verify:{RESET}")
    print(f"    {original_copy}")
    print(f"    {openpyxl_out}")
    print(f"    {offidized_out}")
    print()


if __name__ == "__main__":
    main()
