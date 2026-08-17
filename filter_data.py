import pandas as pd
import numpy as np

def remove_business_applicants(df):
    """
    Remove records where Applicant Name contains business keywords
    Creates two files: 
    1. Cleaned data (business records removed)
    2. Removed records file (for review)
    """
    
    # ============================================================================
    # BUSINESS KEYWORDS (Applicant Name only)
    # ============================================================================
    
    business_keywords = [
        'Ltd', 'Limited', 'PLC', 'LLP', 'CIC',
        'Group', 'Holdings', 'Partnership',
        'Consulting', 'Services', 'Development',
        'School', 'Council', 'Church',
        'Estates', 'Properties', 'Company',
        'Corporation', 'Incorporated',
        'Associates', 'Partners', 'Partnership',
        'Trading', 'Enterprises', 'Ventures',
        'Capital', 'Investments', 'Management',
        'Solutions', 'Systems', 'Technologies',
        'Designs', 'Constructions', 'Builders',
        'Contractors', 'Developers'
    ]
    
    # Additional patterns to check
    business_patterns = [
        'c/o', 'C/O', 'c/', 'C/',  # Care of
        'agent', 'Agent',  # Agent references
        'ltd', 'limited',  # Already in keywords but double-check
    ]
    
    # ============================================================================
    # FUNCTION TO CHECK IF APPLICANT IS BUSINESS
    # ============================================================================
    
    def is_business_applicant(name):
        """
        Check if applicant name indicates a business
        Returns True if business, False if individual
        """
        # Check for NaN or empty
        if pd.isna(name) or str(name).strip() == '':
            return True  # Blank names are excluded
        
        name_str = str(name).lower().strip()
        
        # Check for business patterns first
        for pattern in business_patterns:
            if pattern.lower() in name_str:
                return True
        
        # Check for business keywords
        for keyword in business_keywords:
            if keyword.lower() in name_str:
                return True
        
        # Check for multiple capitalized words (often business names)
        # Only if the name has more than 3 words and looks like a business
        words = name_str.split()
        if len(words) >= 3:
            # Count capitalized words (proper nouns) - simple heuristic
            # If most words are capitalized, it might be a business name
            cap_count = sum(1 for w in words if w[0].isupper() if w)
            if cap_count >= len(words) * 0.7 and len(words) >= 3:
                # Check if it contains common business indicators
                business_indicators = ['inc', 'corp', 'co', 'llc', 'llp']
                if any(indicator in name_str for indicator in business_indicators):
                    return True
        
        return False
    
    # ============================================================================
    # APPLY FILTERING
    # ============================================================================
    
    # Create a copy for safety
    df_clean = df.copy()
    
    # Identify business applicants
    df_clean['Is_Business_Applicant'] = df_clean['Applicant Name'].apply(is_business_applicant)
    
    # Split into removed and kept records
    removed_records = df_clean[df_clean['Is_Business_Applicant'] == True].copy()
    kept_records = df_clean[df_clean['Is_Business_Applicant'] == False].copy()
    
    # Remove the temporary column from kept records
    kept_records = kept_records.drop('Is_Business_Applicant', axis=1)
    
    # ============================================================================
    # ADD REASON FOR REMOVAL
    # ============================================================================
    
    def get_removal_reason(name):
        if pd.isna(name) or str(name).strip() == '':
            return 'Blank Applicant Name'
        
        name_str = str(name).lower()
        
        # Check for specific reasons
        if 'ltd' in name_str:
            return 'Contains "Ltd"'
        if 'limited' in name_str:
            return 'Contains "Limited"'
        if 'c/o' in name_str or 'c/' in name_str:
            return 'C/O Agent - Business Representative'
        if 'agent' in name_str:
            return 'Contains "Agent"'
        if 'group' in name_str:
            return 'Contains "Group"'
        if 'partnership' in name_str:
            return 'Contains "Partnership"'
        if 'consulting' in name_str:
            return 'Contains "Consulting"'
        if 'services' in name_str:
            return 'Contains "Services"'
        if 'development' in name_str:
            return 'Contains "Development"'
        if 'associates' in name_str:
            return 'Contains "Associates"'
        if 'company' in name_str:
            return 'Contains "Company"'
        if 'corporation' in name_str:
            return 'Contains "Corporation"'
        if 'incorporated' in name_str:
            return 'Contains "Incorporated"'
        
        return 'Business Name Detected'
    
    removed_records['Removal_Reason'] = removed_records['Applicant Name'].apply(get_removal_reason)
    
    # ============================================================================
    # GENERATE SUMMARY
    # ============================================================================
    
    print("="*70)
    print("BUSINESS APPLICANT FILTERING RESULTS")
    print("="*70)
    print(f"Total records processed: {len(df)}")
    print(f"Records kept: {len(kept_records)}")
    print(f"Records removed: {len(removed_records)}")
    print(f"Removal rate: {len(removed_records)/len(df)*100:.2f}%")
    
    # Count by removal reason
    print("\nRemoval Reasons Breakdown:")
    print("-"*50)
    reason_counts = removed_records['Removal_Reason'].value_counts()
    for reason, count in reason_counts.items():
        print(f"  {reason}: {count} records")
    
    # Show examples of removed records
    print("\nSample of Removed Records (First 10):")
    print("-"*50)
    sample_columns = ['S.No.', 'Applicant Name', 'Application Type', 'Removal_Reason']
    sample_columns = [col for col in sample_columns if col in removed_records.columns]
    print(removed_records[sample_columns].head(10).to_string(index=False))
    
    # ============================================================================
    # SAVE FILES
    # ============================================================================
    
    # Save kept records
    kept_records.to_csv('applications_cleaned.csv', index=False)
    print(f"\n✓ Cleaned data saved to: applications_cleaned.csv")
    
    # Save removed records with reasons
    removed_records.to_csv('applications_removed_business.csv', index=False)
    print(f"✓ Removed business records saved to: applications_removed_business.csv")
    
    # ============================================================================
    # ADD LETTER NAME COLUMN TO CLEANED DATA
    # ============================================================================
    
    def generate_letter_name(row):
        """
        Generate the appropriate letter name for the applicant
        Only for individual applicants (businesses have been removed)
        """
        applicant_name = row.get('Applicant Name', '')
        
        if pd.isna(applicant_name) or applicant_name == '':
            return 'No Applicant Name'
        
        name_str = str(applicant_name).strip()
        name_parts = name_str.split()
        
        # Check for joint applicants
        if '&' in name_str or ' AND ' in name_str.upper():
            return f'Letter Two - Joint: {name_str}'
        
        # Surname only (one word)
        if len(name_parts) == 1:
            return f'Letter Two - Household: {name_str} Household'
        
        # Check if first name is an initial (e.g., "R Smith")
        first_part = name_parts[0]
        if len(first_part) == 1 or (len(first_part) == 2 and first_part[1] == '.'):
            surname = name_parts[-1] if len(name_parts) > 1 else name_str
            return f'Letter Two - Household: {surname} Household'
        
        # Normal case - use first name only
        return f'Letter Two - Individual: {name_parts[0]}'
    
    def get_salutation(row):
        """Generate the salutation for the letter"""
        applicant_name = row.get('Applicant Name', '')
        
        if pd.isna(applicant_name) or applicant_name == '':
            return 'Dear [No Applicant Name]'
        
        name_str = str(applicant_name).strip()
        
        # Joint applicants
        if '&' in name_str or ' AND ' in name_str.upper():
            return f'Dear {name_str}'
        
        name_parts = name_str.split()
        
        # Surname only
        if len(name_parts) == 1:
            return f'Dear {name_str} Household'
        
        # Initial as first name
        first_part = name_parts[0]
        if len(first_part) == 1 or (len(first_part) == 2 and first_part[1] == '.'):
            surname = name_parts[-1] if len(name_parts) > 1 else name_str
            return f'Dear {surname} Household'
        
        # Normal - use first name
        return f'Dear {name_parts[0]}'
    
    # Add letter columns if records exist
    if len(kept_records) > 0:
        kept_records['Letter_Name'] = kept_records.apply(generate_letter_name, axis=1)
        kept_records['Salutation'] = kept_records.apply(get_salutation, axis=1)
        
        # Save updated file with letter columns
        kept_records.to_csv('applications_cleaned_with_letters.csv', index=False)
        print(f"✓ Cleaned data with letter names saved to: applications_cleaned_with_letters.csv")
        
        # Show letter distribution
        print("\nLetter Name Distribution:")
        print("-"*50)
        print(kept_records['Letter_Name'].value_counts().to_string())
    
    return kept_records, removed_records

# ============================================================================
# EXECUTE THE SCRIPT
# ============================================================================

# Load your CSV file
df = pd.read_csv(r'c:\Users\hp\Downloads\final_results_new.csv')

# Apply business filtering
cleaned_data, removed_data = remove_business_applicants(df)

# ============================================================================
# PREVIEW RESULTS
# ============================================================================

print("\n" + "="*70)
print("PREVIEW OF CLEANED DATA (First 5 Records)")
print("="*70)

# Select columns to preview
preview_columns = ['S.No.', 'Applicant Name', 'Application Type', 'Letter_Name', 'Salutation']
preview_columns = [col for col in preview_columns if col in cleaned_data.columns]

if len(cleaned_data) > 0:
    print(cleaned_data[preview_columns].head(10).to_string(index=False))
else:
    print("No records remain after filtering. Check your data!")

print("\n" + "="*70)
print("PREVIEW OF REMOVED RECORDS (First 5 Records)")
print("="*70)

removed_preview = ['S.No.', 'Applicant Name', 'Application Type', 'Removal_Reason']
removed_preview = [col for col in removed_preview if col in removed_data.columns]

if len(removed_data) > 0:
    print(removed_data[removed_preview].head(10).to_string(index=False))
else:
    print("No records were removed.")

# ============================================================================
# CREATE COMPREHENSIVE SUMMARY REPORT
# ============================================================================

summary_df = pd.DataFrame({
    'Metric': [
        'Total Records',
        'Records Kept (Individuals)',
        'Records Removed (Business Applicants)',
        'Removal Rate',
        'Cleaned Data File',
        'Removed Data File'
    ],
    'Value': [
        len(df),
        len(cleaned_data),
        len(removed_data),
        f"{len(removed_data)/len(df)*100:.2f}%",
        'applications_cleaned_with_letters.csv',
        'applications_removed_business.csv'
    ]
})

print("\n" + "="*70)
print("FINAL SUMMARY")
print("="*70)
print(summary_df.to_string(index=False))