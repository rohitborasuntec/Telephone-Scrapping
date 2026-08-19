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
    def __init__(self, excel_file_path, template_path, api_key):
        """
        Initialize the letter processor.
        
        Args:
            excel_file_path: Path to the Excel file
            template_path: Path to the Word template
            api_key: Stannp API key
        """
        self.excel_file_path = excel_file_path
        self.template_path = template_path
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
    
    def generate_filename(self, reference, neighbour_type):
        """Generate filename based on reference and neighbour type."""
        safe_reference = self.sanitize_filename(reference)
        safe_reference = safe_reference.replace('/', '_')
        return f"{safe_reference}_{neighbour_type}"
    
    def format_address(self, address):
        """
        Format address with proper line breaks.
        Example: "85 OPHIR ROAD, PORTSMOUTH, PO2 9ER"
        becomes:
        "85 OPHIR ROAD,
        PORTSMOUTH,
        PO2 9ER"
        """
        if not address:
            return address
        
        # Split by comma and strip whitespace
        parts = [part.strip() for part in address.split(',')]
        
        # Join with newline
        return ',\n'.join(parts)
    
    def generate_letter(self, reference, property_address, neighbour_address, docx_path):
        """Generate a letter from the template."""
        try:
            os.makedirs(os.path.dirname(docx_path), exist_ok=True)
            
            doc = DocxTemplate(self.template_path)
            
            # Format addresses with proper line breaks
            property_address_formatted = self.format_address(property_address)
            neighbour_address_formatted = self.format_address(neighbour_address)
            
            doc.render({
                "Reference": reference,
                "Address": property_address_formatted,  # Property address with line breaks
                "address1": neighbour_address_formatted,  # Neighbour address with line breaks
            })
            
            doc.save(docx_path)
            print(f"✅ Generated DOCX: {os.path.basename(docx_path)}")
            return docx_path
            
        except Exception as e:
            print(f"❌ Error generating letter: {e}")
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
                pdf_url = response_data.get('data').get('pdf', '').replace('\\', '')
                print(f"✅ Successfully sent to Stannp , URL : ",pdf_url)
                return pdf_url
            else:
                print(f"❌ Error sending to Stannp: {response.status_code}")
                print(f"Response: {response.text}")
                return None
                
        except Exception as e:
            print(f"❌ Error in Stannp API call: {e}")
            return None
    
    def process_neighbour_address(self, reference, property_address, neighbour_address, neighbour_type, output_dir, test_mode=True):
        """Process a single neighbour address."""
        try:
            base_filename = self.generate_filename(reference, neighbour_type)
            
            temp_docx_path = os.path.join(self.temp_dir, f"{base_filename}.docx")
            final_pdf_path = os.path.join(output_dir, f"{base_filename}.pdf")
            
            # Generate DOCX
            docx_file = self.generate_letter(
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
    
    def process_row(self, row, output_dir, test_mode=True):
        """Process a single row from the DataFrame."""
        reference = row['Reference']
        property_address = row['Address']
        
        print(f"\n{'='*60}")
        print(f"📝 Processing {reference}")
        print(f"📍 Property Address: {property_address}")
        print(f"{'='*60}")
        
        pdf_urls = {}
        
        # Process Neighbour 1
        neighbour1_address = row.get('Neighbour 1 Address')
        if pd.notna(neighbour1_address) and neighbour1_address:
            print(f"👤 Processing Neighbour 1: {neighbour1_address}")
            pdf_url = self.process_neighbour_address(
                reference, property_address, neighbour1_address, 'One', output_dir, test_mode
            )
            if pdf_url:
                pdf_urls['Neighbour 1'] = pdf_url
        else:
            print(f"⚠️ No Neighbour 1 address")
        
        # Process Neighbour 2
        neighbour2_address = row.get('Neighbour 2 Address')
        if pd.notna(neighbour2_address) and neighbour2_address:
            print(f"👤 Processing Neighbour 2: {neighbour2_address}")
            pdf_url = self.process_neighbour_address(
                reference, property_address, neighbour2_address, 'Two', output_dir, test_mode
            )
            if pdf_url:
                pdf_urls['Neighbour 2'] = pdf_url
        else:
            print(f"⚠️ No Neighbour 2 address")
        
        return pdf_urls
    
    def cleanup_temp_files(self):
        """Clean up the temporary DOCX directory."""
        try:
            if os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)
                print(f"🗑️ Cleaned up temporary directory: {self.temp_dir}")
        except Exception as e:
            print(f"⚠️ Could not clean up temp directory: {e}")
    
    def process_all_rows(self, output_dir="output", test_mode=True):
        """Process all rows in the DataFrame."""
        os.makedirs(output_dir, exist_ok=True)
        os.makedirs(self.temp_dir, exist_ok=True)
        
        if self.df is None:
            self.load_excel()
        
        if self.df.empty:
            print("❌ No data found in Excel file.")
            return self.df
        
        # Check conversion methods before processing
        self.check_conversion_methods()
        
        # Initialize columns
        if 'Letter HTML Neighbour 1' not in self.df.columns:
            self.df['Letter HTML Neighbour 1'] = None
        if 'Letter HTML Neighbour 2' not in self.df.columns:
            self.df['Letter HTML Neighbour 2'] = None
        
        # Process each row
        total_rows = len(self.df)
        for idx, row in self.df.iterrows():
            print(f"\n📊 Processing row {idx + 1}/{total_rows}")
            pdf_urls = self.process_row(row, output_dir, test_mode)
            
            if 'Neighbour 1' in pdf_urls:
                self.df.at[idx, 'Letter HTML Neighbour 1'] = pdf_urls['Neighbour 1']
            
            if 'Neighbour 2' in pdf_urls:
                self.df.at[idx, 'Letter HTML Neighbour 2'] = pdf_urls['Neighbour 2']

            temp_excel = self.excel_file_path.replace('.xlsx', '_temp.xlsx')
            self.df.to_excel(temp_excel, index=False)
            print(f"\n✅ Updated Excel file saved to: {temp_excel}")
        
        # Clean up
        self.cleanup_temp_files()
        
        # Save updated Excel
        output_excel = self.excel_file_path.replace('.xlsx', '_updated.xlsx')
        self.df.to_excel(output_excel, index=False)
        print(f"\n✅ Updated Excel file saved to: {output_excel}")
        
        return self.df

def main():
    # Load environment variables
    try:
        dot_env = dotenv_values(".env")
        API_KEY = dot_env.get("STANNP_KEY")
    except:
        API_KEY = None
    
    # Configuration
    EXCEL_FILE = r"c:\Users\Rohit\Downloads\U095 Extracted Data.xlsx"
    TEMPLATE_PATH = r"c:\Users\Rohit\Downloads\letter_one.docx"
    
    
    if not API_KEY:
        print("⚠️ STANNP_KEY not found in .env file")
        print("Using fallback API key for testing")
    
    OUTPUT_DIR = "output_pdfs"
    TEST_MODE = True
    
    print("\n" + "="*60)
    print("📧 Letter Processing System")
    print("="*60)
    
    # Check if files exist
    if not os.path.exists(TEMPLATE_PATH):
        print(f"❌ Template file not found: {TEMPLATE_PATH}")
        return
    
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ Excel file not found: {EXCEL_FILE}")
        return
    
    # Initialize and process
    processor = LetterProcessor(EXCEL_FILE, TEMPLATE_PATH, API_KEY)
    
    try:
        updated_df = processor.process_all_rows(OUTPUT_DIR, TEST_MODE)
        print("\n✅ Processing complete!")
        
        # Show results
        print("\n📋 Sample results:")
        sample_cols = ['Reference', 'Total Neighbours', 'Letter HTML Neighbour 1', 'Letter HTML Neighbour 2']
        print(updated_df[sample_cols].head(5).to_string())
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()