import csv
import re
from pathlib import Path
from typing import Dict, List, Optional

class ContactConverter:
    def __init__(self):
        self.vcard_template = """BEGIN:VCARD
VERSION:2.1
N:;{name};;;
FN:{name}
TEL;CELL;PREF:{phone}
END:VCARD"""

    def clean_phone_number(self, phone: str) -> str:
        """
        Clean and format phone number by removing spaces and adding country code if missing.
        """
        # Remove all spaces and non-numeric characters
        cleaned = re.sub(r'\D', '', phone)
        
        # Add default country code (+91 for India) if not present
        if len(cleaned) == 10:
            cleaned = f"+91{cleaned}"
        elif len(cleaned) > 10 and not cleaned.startswith('+'):
            cleaned = f"+{cleaned}"
            
        # Format the number with spaces for readability
        if cleaned.startswith('+91'):
            cleaned = f"{cleaned[:3]} {cleaned[3:8]} {cleaned[8:]}"
        
        return cleaned

    def validate_contact(self, contact: Dict[str, str]) -> Optional[Dict[str, str]]:
        """
        Validate contact data and return None if invalid.
        """
        # Get name and phone, handling possible space in column names
        name = None
        phone = None
        
        # Try to find the name column
        for key in contact:
            if key.strip().lower() == 'name':
                name = contact[key].strip()
                break
                
        # Try to find the phone column
        for key in contact:
            if key.strip().lower() == 'phone':
                phone = contact[key].strip()
                break
        
        if not name or not phone:
            print(f"Invalid contact - missing name or phone: {contact}")
            return None
            
        # Basic validation
        if len(name) < 2:
            print(f"Invalid contact - name too short: {contact}")
            return None
            
        cleaned_phone = self.clean_phone_number(phone)
        if not re.search(r'\d', cleaned_phone):
            print(f"Invalid contact - invalid phone number: {contact}")
            return None
            
        return {'name': name, 'phone': cleaned_phone}

    def read_contacts(self, csv_file: str) -> List[Dict[str, str]]:
        """
        Read and validate contacts from CSV file.
        """
        valid_contacts = []
        invalid_contacts = []
        
        try:
            with open(csv_file, mode='r', encoding='utf-8') as file:
                # Print CSV headers for debugging
                reader = csv.DictReader(file)
                print(f"\nCSV Headers found: {reader.fieldnames}")
                
                for row in reader:
                    contact = self.validate_contact(row)
                    if contact:
                        valid_contacts.append(contact)
                    else:
                        invalid_contacts.append(row)
                        
            print(f"\nProcessed {len(valid_contacts)} valid contacts")
            if invalid_contacts:
                print(f"Found {len(invalid_contacts)} invalid contacts")
                    
            return valid_contacts
            
        except (FileNotFoundError, csv.Error) as e:
            raise Exception(f"Error reading CSV file: {str(e)}")

    def create_vcards(self, contacts: List[Dict[str, str]]) -> str:
        """
        Create vCard content from validated contacts.
        """
        return "\n".join(
            self.vcard_template.format(**contact)
            for contact in contacts
        )

    def convert_to_vcard(self, csv_file: str, output_file: str) -> None:
        """
        Convert CSV contacts to vCard format with error handling.
        """
        try:
            # Ensure input file exists
            if not Path(csv_file).is_file():
                raise FileNotFoundError(f"Input file '{csv_file}' not found")

            # Read and validate contacts
            contacts = self.read_contacts(csv_file)
            
            if not contacts:
                raise ValueError("No valid contacts found in the CSV file")

            # Create vCards content
            vcards_content = self.create_vcards(contacts)

            # Write to output file
            output_path = Path(output_file)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(output_file, mode='w', encoding='utf-8') as file:
                file.write(vcards_content)

            print(f"\nSuccess: Created vCard file '{output_file}' with {len(contacts)} contacts")

        except Exception as e:
            print(f"\nError: {str(e)}")
            raise

def main():
    # Input and output file paths
    input_csv = "Untitled spreadsheet - Sheet1.csv"
    output_vcf = "contacts_rys.vcf"

    # Create converter instance and process files
    converter = ContactConverter()
    converter.convert_to_vcard(input_csv, output_vcf)

if __name__ == "__main__":
    main()

# Input and output file paths
# input_csv = "Untitled spreadsheet - Sheet1.csv"  # Replace with your CSV file name
# output_vcf = "contacts_rys.vcf"  # Output vCard file

# Convert CSV contacts to vCard format
# convert_to_vcard(input_csv, output_vcf)
