"""
CLI entry point for the vCard conversion tool.

Usage examples:
  python -m src.cli convert -i contacts.csv -o out.vcf
  python -m src.cli convert -i contacts.csv --from google-csv --country-code +1
  python -m src.cli convert -i data.txt --from text
  python -m src.cli merge file1.vcf file2.vcf -o merged.vcf --dedupe
  python -m src.cli paste               # reads from stdin
  python -m src.cli info contacts.vcf   # show stats without converting
"""
from __future__ import annotations

import argparse
import datetime
import json
import sys
from pathlib import Path
from typing import List, Optional

from rich.console import Console
from rich.panel import Panel
from rich.progress import track
from rich.prompt import Confirm, Prompt
from rich.table import Table
from rich import print as rprint

from vcard_convert.core.models import Contact, ParseResult, SourceFormat, VCardVersion
from vcard_convert.core.parsers.csv_parser import CSVParser
from vcard_convert.core.parsers.text_parser import FreeTextParser
from vcard_convert.core.parsers.vcf_parser import VCFParser
from vcard_convert.core.writers.vcf_writer import VCFWriter

console = Console()

# ── Format auto-detection ─────────────────────────────────────────────────────

_FORMAT_BY_EXT = {
    ".csv": SourceFormat.CSV,
    ".vcf": SourceFormat.VCF,
    ".vcard": SourceFormat.VCF,
    ".txt": SourceFormat.TEXT,
}

_FORMAT_ALIASES = {
    "csv": SourceFormat.CSV,
    "google-csv": SourceFormat.GOOGLE_CSV,
    "google_csv": SourceFormat.GOOGLE_CSV,
    "vcf": SourceFormat.VCF,
    "vcard": SourceFormat.VCF,
    "text": SourceFormat.TEXT,
    "txt": SourceFormat.TEXT,
}


def detect_format(path: Path, hint: Optional[str]) -> SourceFormat:
    if hint:
        return _FORMAT_ALIASES.get(hint.lower(), SourceFormat.UNKNOWN)
    return _FORMAT_BY_EXT.get(path.suffix.lower(), SourceFormat.UNKNOWN)


def get_parser(fmt: SourceFormat, country_code: str):
    if fmt == SourceFormat.CSV:
        return CSVParser(country_code=country_code)
    if fmt == SourceFormat.GOOGLE_CSV:
        return CSVParser(country_code=country_code, force_google_mode=True)
    if fmt == SourceFormat.VCF:
        return VCFParser(country_code=country_code)
    if fmt == SourceFormat.TEXT:
        return FreeTextParser(country_code=country_code)
    raise ValueError(f"Unsupported source format: {fmt}")


# ── Rich display helpers ───────────────────────────────────────────────────────

def preview_contacts(contacts: List[Contact], max_rows: int = 8) -> None:
    table = Table(title="Preview (first rows)", show_lines=True, highlight=True)
    table.add_column("#", style="dim", width=4)
    table.add_column("Full Name", style="bold cyan")
    table.add_column("Phone(s)")
    table.add_column("Email")
    table.add_column("Org")

    for i, c in enumerate(contacts[:max_rows], start=1):
        conf_marker = " [yellow]⚠[/yellow]" if c.confidence < 1.0 else ""
        table.add_row(
            str(i),
            c.full_name + conf_marker,
            ", ".join(c.phones),
            ", ".join(c.emails),
            c.organization,
        )

    if len(contacts) > max_rows:
        table.add_row("…", f"[dim]+{len(contacts) - max_rows} more[/dim]", "", "", "")

    console.print(table)


def display_result_summary(result: ParseResult) -> None:
    console.print(
        Panel.fit(
            f"[green]✓ {len(result.contacts)} contacts[/green]  "
            f"[red]✗ {len(result.errors)} errors[/red]  "
            f"[dim]source: {result.source_format.value}[/dim]",
            title="Parse Result",
        )
    )
    if result.errors:
        table = Table(title="Skipped / Errors", show_lines=False, style="red")
        table.add_column("Line", width=6)
        table.add_column("Reason")
        table.add_column("Raw", no_wrap=False, max_width=50)
        for e in result.errors[:20]:
            table.add_row(str(e.line_number or ""), e.reason, e.raw[:80])
        if len(result.errors) > 20:
            table.add_row("…", f"+{len(result.errors) - 20} more", "")
        console.print(table)

    low_confidence = [c for c in result.contacts if c.confidence < 1.0]
    if low_confidence:
        console.print(
            f"[yellow]⚠  {len(low_confidence)} contacts have low confidence "
            "(from free-text parsing). Review before use.[/yellow]"
        )


def save_error_report(result: ParseResult, output_dir: Path) -> Optional[Path]:
    if not result.errors:
        return None
    output_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = output_dir / f"error_report_{ts}.json"
    data = [
        {"line": e.line_number, "reason": e.reason, "raw": e.raw}
        for e in result.errors
    ]
    report_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    return report_path


def interactive_fallback_convert(args: argparse.Namespace) -> argparse.Namespace:
    """Prompt for any missing required arguments."""

    if not args.input:
        args.input = Prompt.ask("[bold]Input file[/bold] (CSV, VCF, or text file)")

    input_path = Path(args.input)
    if not input_path.exists():
        console.print(f"[red]Error:[/red] File not found: {args.input}")
        sys.exit(1)

    if not args.format:
        detected = detect_format(input_path, None)
        if detected == SourceFormat.UNKNOWN:
            fmt_choice = Prompt.ask(
                "Could not detect format. Choose",
                choices=["csv", "google-csv", "vcf", "text"],
                default="csv",
            )
            args.format = fmt_choice
        else:
            args.format = detected.value

    if not args.output:
        stem = input_path.stem
        default_out = str(Path(args.output_dir) / f"{stem}.vcf")
        args.output = Prompt.ask("Output file", default=default_out)

    return args


# ── Sub-commands ──────────────────────────────────────────────────────────────

def cmd_convert(args: argparse.Namespace) -> int:
    """Convert a CSV / VCF / text file to vCard."""
    args = interactive_fallback_convert(args)

    input_path = Path(args.input)
    fmt = detect_format(input_path, args.format)

    if fmt == SourceFormat.UNKNOWN:
        console.print(
            "[red]Error:[/red] Cannot determine input format. "
            "Use [bold]--from csv|google-csv|vcf|text[/bold]."
        )
        return 1

    parser = get_parser(fmt, args.country_code)

    with console.status(f"Parsing [bold]{input_path.name}[/bold]…"):
        result = parser.parse_file(input_path)

    display_result_summary(result)

    if not result.contacts:
        console.print("[red]No valid contacts found. Aborting.[/red]")
        return 1

    preview_contacts(result.contacts)

    if not args.yes:
        if not Confirm.ask(
            f"Write [bold]{len(result.contacts)}[/bold] contacts to "
            f"[cyan]{args.output}[/cyan]?",
            default=True,
        ):
            console.print("Aborted.")
            return 0

    version = VCardVersion.V3_0 if getattr(args, "vcf_version", "3.0") == "3.0" else VCardVersion.V2_1
    writer = VCFWriter(version=version, dedupe=args.dedupe)

    output_path = Path(args.output)
    writer.write_file(result.contacts, output_path)
    console.print(f"[green]✓[/green] Wrote [bold]{len(result.contacts)}[/bold] contacts → [cyan]{output_path}[/cyan]")

    # Error report
    if result.errors:
        rp = save_error_report(result, Path(args.output_dir))
        if rp:
            console.print(f"[yellow]⚠[/yellow]  Error report → [dim]{rp}[/dim]")

    return 0


def cmd_merge(args: argparse.Namespace) -> int:
    """Merge multiple VCF files into one."""
    if not args.inputs:
        console.print("[red]Error:[/red] Provide at least one VCF file to merge.")
        return 1

    parser = VCFParser(country_code=args.country_code)
    all_contacts: List[Contact] = []
    has_errors = False

    for path_str in track(args.inputs, description="Reading VCF files…"):
        p = Path(path_str)
        if not p.exists():
            console.print(f"[red]  ✗ Not found:[/red] {p}")
            has_errors = True
            continue
        result = parser.parse_file(p)
        all_contacts.extend(result.contacts)
        if result.errors:
            has_errors = True
            console.print(f"[yellow]  ⚠ {len(result.errors)} errors in {p.name}[/yellow]")

    console.print(f"\nLoaded [green]{len(all_contacts)}[/green] contacts total.")

    if not all_contacts:
        console.print("[red]No valid contacts found. Aborting.[/red]")
        return 1

    output_path = Path(args.output)
    writer = VCFWriter(dedupe=args.dedupe)
    writer.write_file(all_contacts, output_path)

    written = len(all_contacts)
    console.print(f"[green]✓[/green] Merged → [cyan]{output_path}[/cyan] ({written} contacts)")
    return 0


def cmd_paste(args: argparse.Namespace) -> int:
    """Read contact data from stdin (for paste workflows)."""
    console.print(
        Panel(
            "Paste your contact data below, then press [bold]Ctrl+D[/bold] "
            "(Unix) or [bold]Ctrl+Z Enter[/bold] (Windows) when done.",
            title="Paste mode",
            style="cyan",
        )
    )
    try:
        text = sys.stdin.read()
    except KeyboardInterrupt:
        console.print("\nAborted.")
        return 0

    if not text.strip():
        console.print("[red]No input received.[/red]")
        return 1

    parser = FreeTextParser(country_code=args.country_code)
    result = parser.parse_text(text)

    display_result_summary(result)

    if not result.contacts:
        console.print("[red]No valid contacts found.[/red]")
        return 1

    preview_contacts(result.contacts)

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    default_out = str(Path(args.output_dir) / f"pasted_contacts_{ts}.vcf")
    output_str = Prompt.ask("Output file", default=default_out)
    output_path = Path(output_str)

    writer = VCFWriter(dedupe=args.dedupe)
    writer.write_file(result.contacts, output_path)
    console.print(f"[green]✓[/green] Wrote [bold]{len(result.contacts)}[/bold] contacts → [cyan]{output_path}[/cyan]")
    return 0


def cmd_info(args: argparse.Namespace) -> int:
    """Show statistics for a contact file without converting."""
    p = Path(args.input)
    if not p.exists():
        console.print(f"[red]File not found:[/red] {p}")
        return 1

    fmt = detect_format(p, getattr(args, "format", None))
    parser = get_parser(fmt if fmt != SourceFormat.UNKNOWN else SourceFormat.VCF, args.country_code)
    result = parser.parse_file(p)

    display_result_summary(result)
    preview_contacts(result.contacts, max_rows=10)
    return 0


# ── Argument parser setup ─────────────────────────────────────────────────────

def build_arg_parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(
        prog="vcard",
        description="Flexible contact conversion tool (CSV / text / VCF → vCard).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python -m src.cli convert -i contacts.csv -o out.vcf
  python -m src.cli convert -i contacts.csv --from google-csv
  python -m src.cli convert -i data.txt --from text --country-code +1
  python -m src.cli merge a.vcf b.vcf -o merged.vcf --dedupe
  python -m src.cli paste
  python -m src.cli info contacts.vcf
""",
    )

    # Shared flags
    shared = argparse.ArgumentParser(add_help=False)
    shared.add_argument(
        "--country-code", default="+91", metavar="CC",
        help="Default country code (default: +91)",
    )
    shared.add_argument(
        "--output-dir", default="output", metavar="DIR",
        help="Directory for output files and error reports (default: output/)",
    )
    shared.add_argument(
        "--dedupe", action="store_true",
        help="Remove duplicate contacts (by normalised phone number)",
    )
    shared.add_argument(
        "-y", "--yes", action="store_true",
        help="Skip confirmation prompts",
    )

    sub = root.add_subparsers(dest="command", required=True)

    # ── convert ──────────────────────────────────────────────────────────────
    p_convert = sub.add_parser(
        "convert", parents=[shared],
        help="Convert a file to vCard format",
    )
    p_convert.add_argument("-i", "--input", metavar="FILE", help="Input file path")
    p_convert.add_argument(
        "--from", dest="format", metavar="FORMAT",
        choices=["csv", "google-csv", "vcf", "text"],
        help="Source format (auto-detected from extension if omitted)",
    )
    p_convert.add_argument("-o", "--output", metavar="FILE", help="Output .vcf file path")
    p_convert.add_argument(
        "--vcf-version", dest="vcf_version", choices=["2.1", "3.0"], default="3.0",
        help="vCard version to produce (default: 3.0)",
    )
    p_convert.set_defaults(func=cmd_convert)

    # ── merge ─────────────────────────────────────────────────────────────────
    p_merge = sub.add_parser(
        "merge", parents=[shared],
        help="Merge multiple VCF files into one",
    )
    p_merge.add_argument("inputs", nargs="*", metavar="FILE", help="VCF files to merge")
    p_merge.add_argument("-o", "--output", required=False, metavar="FILE", help="Output .vcf file")
    p_merge.set_defaults(func=cmd_merge)

    # ── paste ─────────────────────────────────────────────────────────────────
    p_paste = sub.add_parser(
        "paste", parents=[shared],
        help="Convert pasted free-form text from stdin",
    )
    p_paste.set_defaults(func=cmd_paste)

    # ── info ──────────────────────────────────────────────────────────────────
    p_info = sub.add_parser(
        "info", parents=[shared],
        help="Show stats for a contact file (no conversion)",
    )
    p_info.add_argument("input", metavar="FILE", help="File to inspect")
    p_info.add_argument(
        "--from", dest="format", metavar="FORMAT",
        choices=["csv", "google-csv", "vcf", "text"],
        help="Source format hint",
    )
    p_info.set_defaults(func=cmd_info)

    return root


# ── Main ──────────────────────────────────────────────────────────────────────

def main(argv: Optional[List[str]] = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    # Default output path for merge if not given
    if args.command == "merge" and not getattr(args, "output", None):
        import datetime
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        args.output = str(Path(args.output_dir) / f"merged_{ts}.vcf")

    try:
        return args.func(args)
    except KeyboardInterrupt:
        console.print("\n[dim]Interrupted.[/dim]")
        return 130


if __name__ == "__main__":
    sys.exit(main())
