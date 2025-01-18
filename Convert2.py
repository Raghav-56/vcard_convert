import csv
import re
import logging
import sys
import pandas as pd
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Set
from datetime import datetime
from dataclasses import dataclass
import json
import argparse
from enum import Enum
import shutil
from rich.console import Console
from rich.table import Table
from rich import print as rprint


class VCardVersion(Enum):
    V2_1 = "2.1"
    V3_0 = "3.0"


@dataclass
class Contact:
    name: str
    phone: str
    email: Optional[str] = None
    organization: Optional[str] = None
    title: Optional[str] = None
    address: Optional[str] = None
    notes: Optional[str] = None
    groups: Optional[Set[str]] = None


class ContactConverter:
    def __init__(
        self,
        country_code: str = "+91",
        vcard_version: VCardVersion = VCardVersion.V3_0,
        name_suffix: str = "",
        output_dir: str = "output",
        backup: bool = True,
        preview_limit: int = 5,
    ):
        self.country_code = country_code
        self.vcard_version = vcard_version
        self.name_suffix = name_suffix
        self.output_dir = Path(output_dir)
        self.backup = backup
        self.preview_limit = preview_limit
        self.console = Console()
        self.setup_logging()
        self.setup_vcard_templates()

    def setup_vcard_templates(self):
        """Configure vCard templates for different versions."""
        self.templates = {
            VCardVersion.V2_1: """BEGIN:VCARD
VERSION:2.1
N:;{name}{suffix};;;
FN:{name}{suffix}
TEL;CELL;PREF:{phone}
{email}{org}{title}{adr}{note}END:VCARD""",
            VCardVersion.V3_0: """BEGIN:VCARD
VERSION:3.0
N:{lastname};{firstname};;;
FN:{fullname}{suffix}
TEL;TYPE=CELL:{phone}
{email}{org}{title}{adr}{note}END:VCARD""",
        }

    def setup_logging(self):
        """Configure logging with timestamps and levels."""
        self.output_dir.mkdir(parents=True, exist_ok=True)
        log_file = self.output_dir / "contact_converter.log"

        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s",
            handlers=[logging.FileHandler(log_file), logging.StreamHandler()],
        )
        self.logger = logging

    def preview_csv(self, csv_file: Path) -> None:
        """Preview the CSV file contents."""
        try:
            df = pd.read_csv(csv_file)
            self.console.print("\n[bold cyan]CSV Preview:[/bold cyan]")

            # Display basic information
            self.console.print(f"\nTotal rows: {len(df)}")
            self.console.print(f"Columns: {', '.join(df.columns)}\n")

            # Create a rich table for preview
            table = Table(title=f"First {self.preview_limit} rows")
            for col in df.columns:
                table.add_column(col)

            for _, row in df.head(self.preview_limit).iterrows():
                table.add_row(*[str(val) for val in row])

            self.console.print(table)

        except Exception as e:
            self.logger.error(f"Error previewing CSV: {str(e)}")

    def backup_file(self, file_path: Path) -> None:
        """Create a backup of the file if it exists."""
        if file_path.exists() and self.backup:
            backup_path = (
                file_path.parent
                / f"{file_path.stem}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}{file_path.suffix}"
            )
            shutil.copy2(file_path, backup_path)
            self.logger.info(f"Backup created: {backup_path}")

    def clean_phone_number(self, phone: str) -> str:
        """Clean and format phone number with improved validation."""
        cleaned = re.sub(r"\D", "", phone)

        if len(cleaned) < 10:
            raise ValueError(f"Phone number too short: {phone}")

        if len(cleaned) == 10:
            cleaned = f"{self.country_code} {cleaned}"
        elif len(cleaned) > 10 and not cleaned.startswith("+"):
            cleaned = f"+{cleaned}"

        cleaned = " ".join(cleaned[i : i + 4] for i in range(0, len(cleaned), 4))
        return cleaned.strip()

    def parse_name(self, name: str) -> Tuple[str, str, str]:
        """Parse name into firstname, lastname, and full display name."""
        parts = name.strip().split()
        if len(parts) == 1:
            return parts[0], "", parts[0]
        elif len(parts) >= 2:
            firstname = parts[0]
            lastname = " ".join(parts[1:])
            return firstname, lastname, name.strip()
        return "", "", ""

    def validate_email(self, email: Optional[str]) -> Optional[str]:
        """Validate email format if provided."""
        if not email:
            return None
        email = email.strip()
        if re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", email):
            return email
        self.logger.warning(f"Invalid email format: {email}")
        return None

    def create_vcard_entry(self, contact: Contact) -> str:
        """Create vCard entry based on selected version."""
        firstname, lastname, fullname = self.parse_name(contact.name)

        # Optional fields
        email_field = (
            f"EMAIL{';TYPE=INTERNET' if self.vcard_version == VCardVersion.V3_0 else ''}:{contact.email}\n"
            if contact.email
            else ""
        )
        org_field = f"ORG:{contact.organization}\n" if contact.organization else ""
        title_field = f"TITLE:{contact.title}\n" if contact.title else ""
        adr_field = (
            f"ADR{';TYPE=HOME' if self.vcard_version == VCardVersion.V3_0 else ''}:;;{contact.address};;;;\n"
            if contact.address
            else ""
        )
        note_field = f"NOTE:{contact.notes}\n" if contact.notes else ""

        if self.vcard_version == VCardVersion.V2_1:
            return self.templates[VCardVersion.V2_1].format(
                name=contact.name,
                suffix=self.name_suffix,
                phone=contact.phone,
                email=email_field,
                org=org_field,
                title=title_field,
                adr=adr_field,
                note=note_field,
            )
        else:
            return self.templates[VCardVersion.V3_0].format(
                firstname=firstname,
                lastname=lastname,
                fullname=fullname,
                suffix=self.name_suffix,
                phone=contact.phone,
                email=email_field,
                org=org_field,
                title=title_field,
                adr=adr_field,
                note=note_field,
            )

    def analyze_contacts(self, contacts: List[Contact]) -> None:
        """Analyze contacts and print statistics."""
        total = len(contacts)
        with_email = sum(1 for c in contacts if c.email)
        with_org = sum(1 for c in contacts if c.organization)
        with_title = sum(1 for c in contacts if c.title)
        with_address = sum(1 for c in contacts if c.address)
        with_notes = sum(1 for c in contacts if c.notes)

        stats_table = Table(title="Contact Statistics")
        stats_table.add_column("Metric")
        stats_table.add_column("Count")
        stats_table.add_column("Percentage")

    def validate_contact(self, row: Dict[str, str]) -> Optional[Contact]:
        """Validate and create Contact object from CSV row."""
        try:
            name = next(
                (row[k].strip() for k in row if k.strip().lower() == "name"), None
            )
            phone = next(
                (row[k].strip() for k in row if k.strip().lower() == "phone"), None
            )

            if not name or not phone:
                raise ValueError("Missing required fields: name or phone")

            cleaned_phone = self.clean_phone_number(phone)

            email = self.validate_email(
                next(
                    (row[k].strip() for k in row if k.strip().lower() == "email"), None
                )
            )
            organization = next(
                (
                    row[k].strip()
                    for k in row
                    if k.strip().lower() in ["organization", "org"]
                ),
                None,
            )
            title = next(
                (row[k].strip() for k in row if k.strip().lower() == "title"), None
            )
            address = next(
                (row[k].strip() for k in row if k.strip().lower() == "address"), None
            )
            notes = next(
                (row[k].strip() for k in row if k.strip().lower() in ["notes", "note"]),
                None,
            )

            return Contact(
                name=name,
                phone=cleaned_phone,
                email=email,
                organization=organization,
                title=title,
                address=address,
                notes=notes,
            )

        except Exception as e:
            self.logger.error(f"Error validating contact: {str(e)} - Row: {row}")
            return None

    def read_contacts(self, csv_file: str) -> Tuple[List[Contact], List[Dict]]:
        """Read contacts from CSV with detailed error reporting."""
        valid_contacts = []
        invalid_contacts = []

        try:
            with open(csv_file, mode="r", encoding="utf-8-sig") as file:
                reader = csv.DictReader(file)
                self.logger.info(f"CSV Headers found: {reader.fieldnames}")

                for row_num, row in enumerate(reader, start=2):
                    contact = self.validate_contact(row)
                    if contact:
                        valid_contacts.append(contact)
                    else:
                        invalid_contacts.append(
                            {
                                "row_number": row_num,
                                "data": row,
                            }
                        )

            summary = f"""
Processing Summary:
------------------
Total contacts processed: {len(valid_contacts) + len(invalid_contacts)}
Valid contacts: {len(valid_contacts)}
Invalid contacts: {len(invalid_contacts)}
"""
            print(summary)
            self.logger.info(summary)

            return valid_contacts, invalid_contacts

        except Exception as e:
            self.logger.error(f"Error reading CSV file: {str(e)}")
            raise

    def export_error_report(self, invalid_contacts: List[Dict]) -> None:
        """Export detailed error report for invalid contacts."""
        if not invalid_contacts:
            return

        report_path = (
            self.output_dir
            / f"error_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        )
        with open(report_path, "w", encoding="utf-8") as f:
            json.dump(invalid_contacts, f, indent=2)
        self.logger.info(f"Error report exported to {report_path}")

    def convert_to_vcard(
        self, csv_file: str, output_file: Optional[str] = None
    ) -> None:
        """Convert CSV contacts to vCard format with comprehensive error handling."""
        try:
            # Convert input path to Path object and resolve it
            input_path = Path(csv_file).resolve()

            # Show current directory and available files
            self.console.print(
                f"\n[bold green]Current working directory:[/bold green] {Path.cwd()}"
            )

            # Print current working directory and available CSV files
            current_dir = Path.cwd()
            print(f"\nCurrent working directory: {current_dir}")
            print("\nAvailable CSV files in current directory:")
            csv_files = list(current_dir.glob("*.csv"))
            if csv_files:
                self.console.print("\n[bold green]Available CSV files:[/bold green]")
                for file in csv_files:
                    self.console.print(f"- {file.name}")
            else:
                self.console.print(
                    "\n[yellow]No CSV files found in the current directory[/yellow]"
                )

            self.console.print(f"\n[bold]Attempting to open:[/bold] {input_path}")

            if not input_path.is_file():
                raise FileNotFoundError(
                    f"Input file '{csv_file}' not found.\n"
                    f"Full path attempted: {input_path}\n"
                    "Please check the filename and ensure it's in the correct directory."
                )

            # Preview CSV contents
            self.preview_csv(input_path)

            # Generate output filename if not provided
            if output_file is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_file = self.output_dir / f"contacts_{timestamp}.vcf"
            else:
                output_file = Path(output_file)

            # Create backup if enabled
            self.backup_file(output_file)

            # Read and validate contacts
            contacts, invalid_contacts = self.read_contacts(str(input_path))

            if not contacts:
                raise ValueError("No valid contacts found in the CSV file")

            # Analyze contacts
            self.analyze_contacts(contacts)

            # Create vCards content
            vcards_content = "\n".join(
                self.create_vcard_entry(contact) for contact in contacts
            )

            # Ensure output directory exists
            output_file.parent.mkdir(parents=True, exist_ok=True)

            # Write vCard file
            with open(output_file, mode="w", encoding="utf-8") as file:
                file.write(vcards_content)

            # Export error report if there were any invalid contacts
            self.export_error_report(invalid_contacts)

            success_msg = f"""
[bold green]Success![/bold green] 
---------
vCard file created: {output_file}
Total contacts: {len(contacts)}
vCard version: {self.vcard_version.value}
Name suffix added: {self.name_suffix if self.name_suffix else 'None'}
"""
            self.console.print(success_msg)
            self.logger.info(success_msg)

        except Exception as e:
            error_msg = f"Error during conversion: {str(e)}"
            self.console.print(f"\n[bold red]Error:[/bold red] {error_msg}")
            self.logger.error(error_msg)
            raise


def main():
    parser = argparse.ArgumentParser(
        description="Convert CSV contacts to vCard format",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    python convert.py "contacts.csv"
    python convert.py "contacts.csv" --output "my_contacts.vcf" --version 2.1
    python convert.py "contacts.csv" --suffix " (Work)" --country-code +1
        """,
    )
    parser.add_argument("input_csv", help="Input CSV file path")
    parser.add_argument("-o", "--output", help="Output vCard file path (optional)")
    parser.add_argument(
        "--country-code", default="+91", help="Default country code (e.g., +91)"
    )
    parser.add_argument(
        "--version", choices=["2.1", "3.0"], default="3.0", help="vCard version"
    )
    parser.add_argument(
        "--suffix", default="", help="Suffix to add to all contact names"
    )
    parser.add_argument(
        "--output-dir", default="output", help="Directory for output files"
    )
    parser.add_argument(
        "--list-files", action="store_true", help="List available CSV files and exit"
    )
    parser.add_argument(
        "--no-backup", action="store_true", help="Disable automatic backup"
    )
    parser.add_argument(
        "--preview-limit", type=int, default=5, help="Number of rows to preview"
    )
    parser.add_argument(
        "--quiet", action="store_true", help="Suppress non-error output"
    )

    args = parser.parse_args()

    try:
        if args.list_files:
            console = Console()
            console.print("\n[bold]Available CSV files in current directory:[/bold]")
            csv_files = list(Path.cwd().glob("*.csv"))
            if csv_files:
                for file in csv_files:
                    console.print(f"- {file.name}")
            else:
                console.print(
                    "[yellow]No CSV files found in the current directory[/yellow]"
                )
            sys.exit(0)

        converter = ContactConverter(
            country_code=args.country_code,
            vcard_version=VCardVersion(args.version),
            name_suffix=args.suffix,
            output_dir=args.output_dir,
            backup=not args.no_backup,
            preview_limit=args.preview_limit,
        )

        converter.convert_to_vcard(args.input_csv, args.output)

    except Exception as e:
        logging.error(f"Program terminated with error: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
