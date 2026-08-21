import pandas as pd
from docxtpl import DocxTemplate
import requests
import base64
import os
import re
import subprocess
import sys
import shutil
from pathlib import Path
from dotenv import dotenv_values
import time

class LetterProcessor:
    def __init__(self, excel_file_path, template_path1, template_path2, api_key):
        """
        Initialize the letter processor.
        
        Args:
            excel_file_path: Path to the Excel file
            template_path1: Path to the Letter One Word template
            template_path2: Path to the Letter Two Word template
            api_key: Stannp API key
        """
        self.excel_file_path = excel_file_path
        self.template_path1 = template_path1  # Letter One template
        self.template_path2 = template_path2  # Letter Two template
        self.api_key = api_key
        self.df = None
        self.temp_dir = "temp_docx"
        self.conversion_method = None  # Will be set after checking dependencies
        
    def load_excel(self):
        """Load the Excel file into a DataFrame."""
        self.df = pd.read_excel(self.excel_file_path)
        print(f"Loaded {len(self.df)} rows from Excel file")
        return self.df
    
    def sanitize_filename(self, filename):
        """Sanitize filename to remove invalid characters."""
        invalid_chars = r'[<>:"/\\|?*]'
        sanitized = re.sub(invalid_chars, '_', filename)
        sanitized = re.sub(r'_+', '_', sanitized)
        sanitized = sanitized.strip('_')
        sanitized = sanitized.replace(' ', '_')
        if len(sanitized) > 200:
            sanitized = sanitized[:200]
        return sanitized
    
    def generate_filename(self, reference, letter_type, neighbour_type=None):
        """
        Generate filename based on reference, letter type and neighbour type.
        
        Args:
            reference: Planning reference
            letter_type: 'One' or 'Two'
            neighbour_type: 'One', 'Two' or None (for applicant)
        """
        safe_reference = self.sanitize_filename(reference)
        safe_reference = safe_reference.replace('/', '_')
        
        if letter_type == 'One' and neighbour_type:
            return f"{safe_reference}_Letter{letter_type}_Neighbour{neighbour_type}"
        elif letter_type == 'Two':
            return f"{safe_reference}_Letter{letter_type}_Applicant"
        else:
            return f"{safe_reference}_Letter{letter_type}"
    
    def format_address(self, address):
        """
        Format address with proper line breaks.
        Example: "85 OPHIR ROAD, PORTSMOUTH, PO2 9ER"
        becomes:
        "85 OPHIR ROAD,\nPORTSMOUTH,\nPO2 9ER"
        """
        if not address:
            return address
        
        # Split by comma and strip whitespace
        parts = [part.strip() for part in address.split(',')]
        
        # Join with newline
        return ',\n'.join(parts)
    
    def generate_letter_one(self, reference, property_address, neighbour_address, docx_path):
        """
        Generate Letter One (Neighbour Letter) from the template.
        """
        try:
            os.makedirs(os.path.dirname(docx_path), exist_ok=True)
            
            doc = DocxTemplate(self.template_path1)
            
            # Format addresses with proper line breaks
            property_address_formatted = self.format_address(property_address)
            neighbour_address_formatted = self.format_address(neighbour_address)
            doc.render({
                "Reference": reference,
                "Address": property_address,  # Property address with line breaks
                "address1": neighbour_address_formatted,  # Neighbour address with line breaks
            })
            
            doc.save(docx_path)
            print(f"✅ Generated Letter One DOCX: {os.path.basename(docx_path)}")
            return docx_path
            
        except Exception as e:
            print(f"❌ Error generating Letter One: {e}")
            return None
    
    def generate_letter_two(self, reference, property_address, applicant_first_name, docx_path):
        """
        Generate Letter Two (Applicant Letter) from the template.
        """
        try:
            os.makedirs(os.path.dirname(docx_path), exist_ok=True)
            
            doc = DocxTemplate(self.template_path2)
            
            # Format address with proper line breaks (SAME AS LETTER ONE)
            property_address_formatted = self.format_address(property_address)
            
            doc.render({
                "Applicant_First_Name": applicant_first_name + ",",
                "Reference": reference,
                "address1": property_address_formatted,  # Property address with line breaks (NOW FORMATTED)
            })
            
            doc.save(docx_path)
            print(f"✅ Generated Letter Two DOCX: {os.path.basename(docx_path)}")
            return docx_path
            
        except Exception as e:
            print(f"❌ Error generating Letter Two: {e}")
            return None
    
    def check_conversion_methods(self):
        """Check which conversion methods are available."""
        methods = []
        
        # Check for docx2pdf (requires Microsoft Word)
        try:
            from docx2pdf import convert
            methods.append('docx2pdf')
            print("✅ docx2pdf available (requires Microsoft Word)")
        except ImportError:
            print("⚠️ docx2pdf not installed")
        
        # Check for LibreOffice
        libreoffice_path = self.find_libreoffice()
        if libreoffice_path:
            methods.append('libreoffice')
            print(f"✅ LibreOffice found at: {libreoffice_path}")
        
        if not methods:
            print("❌ No conversion methods found!")
            print("Please install one of the following:")
            print("1. Microsoft Word + docx2pdf: pip install docx2pdf")
            print("2. LibreOffice: https://www.libreoffice.org/download/download/")
        
        return methods
    
    def find_libreoffice(self):
        """Find LibreOffice installation path on Windows."""
        if sys.platform == "win32":
            possible_paths = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                r"C:\Program Files\LibreOffice\program\soffice.bin",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.bin",
            ]
            for path in possible_paths:
                if os.path.exists(path):
                    return path
            
            # Check if soffice is in PATH
            try:
                result = subprocess.run(["where", "soffice"], capture_output=True, text=True)
                if result.returncode == 0:
                    return result.stdout.strip().split('\n')[0]
            except:
                pass
        else:
            # Linux/Mac
            try:
                result = subprocess.run(["which", "soffice"], capture_output=True, text=True)
                if result.returncode == 0:
                    return result.stdout.strip()
            except:
                pass
        return None
    
    def convert_docx_to_pdf_docx2pdf(self, docx_path, pdf_path):
        """Convert using docx2pdf (requires Microsoft Word)."""
        try:
            from docx2pdf import convert
            convert(docx_path, pdf_path)
            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                return True
        except Exception as e:
            print(f"⚠️ docx2pdf conversion failed: {e}")
        return False
    
    def convert_docx_to_pdf_libreoffice(self, docx_path, pdf_path):
        """Convert using LibreOffice."""
        try:
            libreoffice_path = self.find_libreoffice()
            if not libreoffice_path:
                return False
            
            # Ensure output directory exists
            os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
            
            # Convert using LibreOffice
            cmd = [
                libreoffice_path,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", os.path.dirname(pdf_path),
                docx_path
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            
            # LibreOffice saves with same name but .pdf extension in the outdir
            generated_pdf = os.path.join(os.path.dirname(pdf_path), 
                                        os.path.basename(docx_path).replace('.docx', '.pdf'))
            
            if os.path.exists(generated_pdf) and os.path.getsize(generated_pdf) > 0:
                if generated_pdf != pdf_path:
                    shutil.move(generated_pdf, pdf_path)
                return True
            
            print(f"❌ LibreOffice conversion failed: {result.stderr}")
            return False
            
        except subprocess.TimeoutExpired:
            print("❌ LibreOffice conversion timeout")
            return False
        except Exception as e:
            print(f"❌ LibreOffice conversion error: {e}")
            return False
    
    def convert_docx_to_pdf(self, docx_path, pdf_path):
        """
        Convert DOCX to PDF using available methods.
        """
        # If we haven't checked conversion methods yet
        if self.conversion_method is None:
            methods = self.check_conversion_methods()
            if not methods:
                print("❌ No conversion methods available!")
                print("Please install one of the following:")
                print("1. Microsoft Word + docx2pdf: pip install docx2pdf")
                print("2. LibreOffice: https://www.libreoffice.org/download/download/")
                return None
            
            # Prefer docx2pdf if available (faster on Windows with Word)
            if 'docx2pdf' in methods:
                self.conversion_method = 'docx2pdf'
            else:
                self.conversion_method = 'libreoffice'
            
            print(f"✅ Using conversion method: {self.conversion_method}")
        
        # Try the selected method first
        if self.conversion_method == 'docx2pdf':
            print(f"🔄 Converting using docx2pdf (Microsoft Word)...")
            if self.convert_docx_to_pdf_docx2pdf(docx_path, pdf_path):
                print(f"✅ Converted to PDF using docx2pdf")
                return pdf_path
        
        # Try LibreOffice
        if self.conversion_method == 'libreoffice' or self.conversion_method == 'docx2pdf':
            print(f"🔄 Converting using LibreOffice...")
            if self.convert_docx_to_pdf_libreoffice(docx_path, pdf_path):
                print(f"✅ Converted to PDF using LibreOffice")
                return pdf_path
        
        # If both failed, try the other method if available
        if self.conversion_method == 'docx2pdf':
            print(f"🔄 Trying LibreOffice as fallback...")
            if self.convert_docx_to_pdf_libreoffice(docx_path, pdf_path):
                print(f"✅ Converted to PDF using LibreOffice")
                return pdf_path
        elif self.conversion_method == 'libreoffice':
            print(f"🔄 Trying docx2pdf as fallback...")
            if self.convert_docx_to_pdf_docx2pdf(docx_path, pdf_path):
                print(f"✅ Converted to PDF using docx2pdf")
                return pdf_path
        
        print(f"❌ All conversion methods failed!")
        return None
    
    def send_to_stannp(self, pdf_path, reference, test_mode=True):
        """Send the PDF to Stannp API."""
        try:
            if not os.path.exists(pdf_path):
                print(f"❌ PDF file not found: {pdf_path}")
                return None
            
            file_size = os.path.getsize(pdf_path)
            if file_size == 0:
                print(f"❌ PDF file is empty: {pdf_path}")
                return None
            
            print(f"📤 Sending PDF to Stannp API...")
            
            url = "https://api-eu1.stannp.com/v1/letters/post"
            
            auth = base64.b64encode(f"{self.api_key}:".encode()).decode()
            
            payload = {
                "test": "1" if test_mode else "0",
                "country": "GB",
                "size": "A4",
                "duplex": "true"
            }
            
            with open(pdf_path, "rb") as pdf_file:
                files = {
                    "pdf": (
                        os.path.basename(pdf_path),
                        pdf_file,
                        "application/pdf"
                    )
                }
                
                headers = {
                    "Authorization": f"Basic {auth}"
                }
                
                response = requests.post(
                    url,
                    headers=headers,
                    data=payload,
                    files=files,
                    timeout=60
                )

            if response.status_code == 200:
                response_data = response.json()
                pdf_url = response_data.get('data', {}).get('pdf', '').replace('\\', '')
                print(f"✅ Successfully sent to Stannp, URL: {pdf_url}")
                return pdf_url
            else:
                print(f"❌ Error sending to Stannp: {response.status_code}")
                print(f"Response: {response.text}")
                return None
                
        except Exception as e:
            print(f"❌ Error in Stannp API call: {e}")
            return None
    
    def process_neighbour_address(self, reference, property_address, neighbour_address, 
                                   neighbour_type, output_dir, test_mode=True):
        """
        Process a single neighbour address (Letter One).
        """
        try:
            base_filename = self.generate_filename(reference, 'One', neighbour_type)
            
            temp_docx_path = os.path.join(self.temp_dir, f"{base_filename}.docx")
            final_pdf_path = os.path.join(output_dir, f"{base_filename}.pdf")
            
            # Generate Letter One DOCX
            docx_file = self.generate_letter_one(
                reference, 
                property_address,
                neighbour_address,
                temp_docx_path
            )
            
            if not docx_file:
                return None
            
            # Convert to PDF
            pdf_path = self.convert_docx_to_pdf(docx_file, final_pdf_path)
            
            # Clean up temporary DOCX
            try:
                if os.path.exists(docx_file):
                    os.remove(docx_file)
            except:
                pass
            
            if not pdf_path:
                return None
            
            # Send to Stannp
            pdf_url = self.send_to_stannp(pdf_path, reference, test_mode)
            
            return pdf_url
            
        except Exception as e:
            print(f"❌ Error processing neighbour {neighbour_type}: {e}")
            return None
    
    def process_applicant_letter(self, reference, property_address, applicant_name, 
                                   output_dir, test_mode=True):
        """
        Process applicant letter (Letter Two).
        """
        try:
            # Extract first name from applicant name
            first_name = self.extract_first_name(applicant_name)
            
            base_filename = self.generate_filename(reference, 'Two')
            
            temp_docx_path = os.path.join(self.temp_dir, f"{base_filename}.docx")
            final_pdf_path = os.path.join(output_dir, f"{base_filename}.pdf")
            
            # Generate Letter Two DOCX
            docx_file = self.generate_letter_two(
                reference,
                property_address,
                first_name,
                temp_docx_path
            )
            
            if not docx_file:
                return None
            
            # Convert to PDF
            pdf_path = self.convert_docx_to_pdf(docx_file, final_pdf_path)
            
            # Clean up temporary DOCX
            try:
                if os.path.exists(docx_file):
                    os.remove(docx_file)
            except:
                pass
            
            if not pdf_path:
                return None
            
            # Send to Stannp
            pdf_url = self.send_to_stannp(pdf_path, reference, test_mode)
            
            return pdf_url
            
        except Exception as e:
            print(f"❌ Error processing applicant letter: {e}")
            return None
    
def extract_first_name(self, applicant_name):
    """
    Extract first name from applicant name.
    Handles various name formats:
    - "Mr P Steele" -> "Steele Household"
    - "Mr. Chris Banks & Mrs. Claire Tyler" -> "Mr. Chris Banks & Mrs. Claire Tyler" (Joint)
    - "Smith" -> "Smith Household"
    - "R Smith" -> "Smith Household"
    - "Mr Chris Banks" -> "Chris" (Individual)
    - "Mr And Mrs Pinkham" -> "Mr And Mrs Pinkham" (Couple with only surname)
    - "Mr & Mrs Hillcoat" -> "Mr & Mrs Hillcoat" (Couple with only surname)
    - "Mr and Mrs Collins" -> "Mr and Mrs Collins" (Couple with only surname)
    """
    if pd.isna(applicant_name) or str(applicant_name).strip() == '':
        return 'Applicant'
    
    name_str = str(applicant_name).strip()
    
    # List of business/organisation names to keep as-is
    businesses = [
        "Nationwide Building Society",
        "The Shere Surgery and Dispensary",
        "Allianz",
        "Owner"
    ]
    
    if name_str in businesses:
        return name_str
    
    # Check for "Mr & Mrs" or "Mr and Mrs" patterns (case insensitive)
    # This must be checked BEFORE the general '&' or 'and' detection
    mr_mrs_pattern = r'(?i)^(mr\s*&\s*mrs|mr\s+and\s+mrs)'
    mr_mrs_match = re.match(mr_mrs_pattern, name_str)
    
    if mr_mrs_match:
        # Extract the rest after Mr & Mrs
        rest = name_str[mr_mrs_match.end():].strip()
        
        # Check if rest contains first names (e.g., "Simon & Mary O'Brien" or "Joe and Gillian Allen")
        # If there are 2 or more words after, or contains '&' or 'and', it has first names
        rest_parts = rest.split()
        
        # Check for patterns like "Simon & Mary O'Brien" or "Joe and Gillian Allen"
        if len(rest_parts) >= 3:  # e.g., "Simon & Mary O'Brien" has 3+ words
            return name_str  # Keep full as-is
        elif '&' in rest or ' and ' in rest.lower():
            return name_str  # Keep full as-is
        elif len(rest_parts) >= 2 and rest_parts[0][0].isupper() and rest_parts[1][0].isupper():
            # Two capitalized words = first names (e.g., "Phil and Pam" or "Simon Mary")
            # But check if it's actually a first and last name
            # If it's only 2 words and both are capitalized, it could be "Joe Allen" (first + last)
            # We'll keep it as-is to be safe
            return name_str
        else:
            # Only surname - keep as "Mr & Mrs [Surname]" or "Mr and Mrs [Surname]"
            # Normalize the "and" to "&" for consistency if needed
            return name_str
    
    # Check for joint applicants with '&' or 'and' (but not Mr & Mrs pattern)
    if '&' in name_str or ' AND ' in name_str.upper():
        # Check if it has full names (e.g., "Charles and Amanda Yaxley")
        # Keep full as-is
        return name_str
    
    # Check for titles (Mr, Mrs, Miss, Ms, Dr, Rev, etc.)
    titles = ['mr', 'mrs', 'ms', 'dr', 'prof', 'rev', 'sir', 'lord', 'lady', 'miss', 'mr.', 'mrs.', 'ms.', 'dr.', 'prof.', 'rev.']
    name_parts = name_str.split()
    
    # Check if first part is a title
    first_part_lower = name_parts[0].lower().replace('.', '')
    
    if first_part_lower in titles:
        # Has title - check the rest
        if len(name_parts) == 1:
            # Just title - return as-is
            return name_str
        else:
            rest = ' '.join(name_parts[1:])
            rest_parts = rest.split()
            
            if len(rest_parts) == 0:
                return name_str
            elif len(rest_parts) == 1:
                # Only surname (e.g., "Mr Barfoot")
                return f"{rest_parts[0]} Household"
            else:
                # Check if first part of rest is an initial (e.g., "P Steele")
                first_rest = rest_parts[0]
                if len(first_rest) <= 2 and (first_rest.isupper() or first_rest.endswith('.')):
                    # It's an initial - surname takes precedence
                    surname = rest_parts[-1].strip('.,')
                    return f"{surname} Household"
                else:
                    # Full first name - use first name only (e.g., "Mr Chris Banks" -> "Chris")
                    return rest_parts[0].strip('.,')
    
    # No title - handle other cases
    if len(name_parts) == 1:
        # Single name - it's a surname
        return f"{name_str} Household"
    elif len(name_parts) == 2:
        # Two names - check if first is an initial
        first_part = name_parts[0]
        if len(first_part) <= 2 and (first_part.isupper() or first_part.endswith('.')):
            # Initial + Surname (e.g., "R Smith")
            return f"{name_parts[1]} Household"
        else:
            # First name + Surname (e.g., "Rebecca Clewley")
            return name_parts[0].strip('.,')
    else:
        # Three or more names - likely has first and last name
        # Check if first part is an initial
        first_part = name_parts[0]
        if len(first_part) <= 2 and (first_part.isupper() or first_part.endswith('.')):
            # Initial + rest - surname takes precedence
            return f"{name_parts[-1]} Household"
        else:
            # Use the first part as first name
            return name_parts[0].strip('.,')
    
    def process_row(self, row, output_dir, test_mode=True, process_letter_one=True, process_letter_two=True):
        """
        Process a single row from the DataFrame.
        
        Args:
            row: DataFrame row
            output_dir: Output directory
            test_mode: Whether to run in test mode
            process_letter_one: Whether to process Letter One (neighbour letters)
            process_letter_two: Whether to process Letter Two (applicant letter)
        """
        reference = row['Reference']
        property_address = row['Address']
        applicant_name = row.get('Letter Name', '')
        
        print(f"\n{'='*60}")
        print(f"📝 Processing {reference}")
        print(f"📍 Property Address: {property_address}")
        print(f"👤 Applicant: {applicant_name}")
        print(f"{'='*60}")
        
        pdf_urls = {}
        
        # Process Neighbour 1 (Letter One) - SKIP if not processing Letter One
        if process_letter_one:
            neighbour1_address = row.get('Neighbour 1 Address')
            if pd.notna(neighbour1_address) and neighbour1_address:
                print(f"👤 Processing Neighbour 1: {neighbour1_address}")
                pdf_url = self.process_neighbour_address(
                    reference, property_address, neighbour1_address, 'One', output_dir, test_mode
                )
                if pdf_url:
                    pdf_urls['Letter One Neighbour 1'] = pdf_url
            else:
                print(f"⚠️ No Neighbour 1 address")
            
            # Process Neighbour 2 (Letter One)
            neighbour2_address = row.get('Neighbour 2 Address')
            if pd.notna(neighbour2_address) and neighbour2_address:
                print(f"👤 Processing Neighbour 2: {neighbour2_address}")
                pdf_url = self.process_neighbour_address(
                    reference, property_address, neighbour2_address, 'Two', output_dir, test_mode
                )
                if pdf_url:
                    pdf_urls['Letter One Neighbour 2'] = pdf_url
            else:
                print(f"⚠️ No Neighbour 2 address")
        else:
            print(f"⏭️ Skipping Letter One (neighbour letters)")
        
        # Process Applicant Letter (Letter Two)
        if process_letter_two:
            if pd.notna(applicant_name) and applicant_name:
                print(f"👤 Processing Applicant Letter: {applicant_name}")
                pdf_url = self.process_applicant_letter(
                    reference, property_address, applicant_name, output_dir, test_mode
                )
                if pdf_url:
                    pdf_urls['Letter Two Applicant'] = pdf_url
            else:
                print(f"⚠️ No Applicant name")
        else:
            print(f"⏭️ Skipping Letter Two (applicant letter)")
        
        return pdf_urls
    
    def cleanup_temp_files(self):
        """Clean up the temporary DOCX directory."""
        try:
            if os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)
                print(f"🗑️ Cleaned up temporary directory: {self.temp_dir}")
        except Exception as e:
            print(f"⚠️ Could not clean up temp directory: {e}")
    
    def process_all_rows(self, output_dir="output", test_mode=True, process_letter_one=True, process_letter_two=True):
        """
        Process all rows in the DataFrame.
        
        Args:
            output_dir: Output directory for PDFs
            test_mode: Whether to run in test mode
            process_letter_one: Whether to process Letter One (neighbour letters)
            process_letter_two: Whether to process Letter Two (applicant letter)
        """
        os.makedirs(output_dir, exist_ok=True)
        os.makedirs(self.temp_dir, exist_ok=True)
        
        if self.df is None:
            self.load_excel()
        
        if self.df.empty:
            print("❌ No data found in Excel file.")
            return self.df
        
        print(f"\n📋 Processing settings:")
        print(f"   Letter One (Neighbour): {'✅' if process_letter_one else '❌'}")
        print(f"   Letter Two (Applicant): {'✅' if process_letter_two else '❌'}")
        
        # Check conversion methods before processing
        self.check_conversion_methods()
        
        # Initialize columns for Letter One if needed
        if process_letter_one:
            if 'Letter One HTML Neighbour 1' not in self.df.columns:
                self.df['Letter One HTML Neighbour 1'] = None
            if 'Letter One HTML Neighbour 2' not in self.df.columns:
                self.df['Letter One HTML Neighbour 2'] = None
        
        # Initialize columns for Letter Two if needed
        if process_letter_two:
            if 'Letter Two HTML' not in self.df.columns:
                self.df['Letter Two HTML'] = None
        
        # Process each row
        list_of_index = [68,23,97,114,163,145,194,199,221,252,39]
        total_rows = len(self.df)
        for idx, row in self.df.iterrows():
            if idx not in list_of_index:
                continue
            print(f"\n📊 Processing row {idx + 1}/{total_rows}")
            pdf_urls = self.process_row(row, output_dir, test_mode, process_letter_one, process_letter_two)
            
            # Update DataFrame with URLs
            if 'Letter One Neighbour 1' in pdf_urls:
                self.df.at[idx, 'Letter One HTML Neighbour 1'] = pdf_urls['Letter One Neighbour 1']
            
            if 'Letter One Neighbour 2' in pdf_urls:
                self.df.at[idx, 'Letter One HTML Neighbour 2'] = pdf_urls['Letter One Neighbour 2']
            
            if 'Letter Two Applicant' in pdf_urls:
                self.df.at[idx, 'Letter Two HTML'] = pdf_urls['Letter Two Applicant']

            # Save temporary progress
            temp_excel = self.excel_file_path.replace('.xlsx', '_temp.xlsx')
            self.df.to_excel(temp_excel, index=False)
            print(f"\n💾 Updated Excel file saved to: {temp_excel}")
        
        # Clean up
        self.cleanup_temp_files()
        
        # Save final updated Excel
        output_excel = self.excel_file_path.replace('.xlsx', '_updated.xlsx')
        self.df.to_excel(output_excel, index=False)
        print(f"\n✅ Final updated Excel file saved to: {output_excel}")
        
        return self.df

def main():
    # Load environment variables
    try:
        dot_env = dotenv_values(".env")
        API_KEY = dot_env.get("STANNP_KEY")
    except:
        API_KEY = None
    
    # Configuration
    EXCEL_FILE = r"c:\Users\Rohit\Downloads\Data (1).xlsx"
    TEMPLATE_PATH1 = r"c:\Users\Rohit\Downloads\letter_one.docx"
    TEMPLATE_PATH2 = r"c:\Users\Rohit\Documents\letter two final draft 09.08.2026.docx"
    
    if not API_KEY:
        print("⚠️ STANNP_KEY not found in .env file")
        print("Using fallback API key for testing")
        API_KEY = "YOUR_FALLBACK_API_KEY"
    
    OUTPUT_DIR = "output_pdfs"
    TEST_MODE = True
    
    print("\n" + "="*60)
    print("📧 Letter Processing System")
    print("="*60)
    print(f"📂 Excel File: {EXCEL_FILE}")
    print(f"📄 Template 1 (Letter One): {TEMPLATE_PATH1}")
    print(f"📄 Template 2 (Letter Two): {TEMPLATE_PATH2}")
    print(f"📁 Output Directory: {OUTPUT_DIR}")
    print(f"🧪 Test Mode: {TEST_MODE}")
    print("="*60)
    
    # Check if files exist
    if not os.path.exists(TEMPLATE_PATH1):
        print(f"❌ Template 1 not found: {TEMPLATE_PATH1}")
        return
    
    if not os.path.exists(TEMPLATE_PATH2):
        print(f"❌ Template 2 not found: {TEMPLATE_PATH2}")
        return
    
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ Excel file not found: {EXCEL_FILE}")
        return
    
    # Initialize and process
    processor = LetterProcessor(EXCEL_FILE, TEMPLATE_PATH1, TEMPLATE_PATH2, API_KEY)
    
    try:
        # MODIFIED: Set process_letter_one=False, process_letter_two=True
        # This will ONLY process Letter Two (applicant letters)
        updated_df = processor.process_all_rows(
            OUTPUT_DIR, 
            TEST_MODE,
            process_letter_one=True,
            process_letter_two=True 
        )
        
        print("\n" + "="*60)
        print("✅ Processing complete!")
        print("="*60)
        
        # Show results
        print("\n📋 Sample results:")
        sample_cols = ['Reference', 'Total Neighbours', 
                       'Letter One HTML Neighbour 1', 'Letter One HTML Neighbour 2',
                       'Letter Two HTML']
        existing_cols = [col for col in sample_cols if col in updated_df.columns]
        print(updated_df[existing_cols].head(5).to_string())
        
        # Print statistics
        print("\n📊 Statistics:")
        total_rows = len(updated_df)
        letter_one_count = updated_df['Letter One HTML Neighbour 1'].notna().sum() + \
                           updated_df['Letter One HTML Neighbour 2'].notna().sum()
        letter_two_count = updated_df['Letter Two HTML'].notna().sum()
        
        print(f"   Total rows processed: {total_rows}")
        print(f"   Letter One sent: {letter_one_count}")
        print(f"   Letter Two sent: {letter_two_count}")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()