import pandas as pd
import numpy as np

def merge_with_neighbours(df1, df2):
    """
    Merge df1 and df2 based on Reference field
    Creates final output with neighbours in separate columns
    """
    
    # ============================================================================
    # STEP 1: CLEAN AND PREPARE DATA
    # ============================================================================
    
    # Create copies to avoid modifying originals
    df1_clean = df1.copy()
    df2_clean = df2.copy()
    
    # Clean Reference fields for matching
    df1_clean['Reference_Clean'] = df1_clean['Reference'].astype(str).str.strip().str.upper()
    df2_clean['Reference_Clean'] = df2_clean['Reference Id'].astype(str).str.strip().str.upper()
    
    # ============================================================================
    # STEP 2: MERGE BASED ON REFERENCE
    # ============================================================================
    
    # First, merge to get all matches
    merged_df = pd.merge(
        df1_clean,
        df2_clean,
        on='Reference_Clean',
        how='left'
    )
    
    print(f"Total records after merge: {len(merged_df)}")
    print(f"Records with neighbours: {merged_df['Neighbour Address'].notna().sum()}")
    
    # ============================================================================
    # STEP 3: IDENTIFY NEIGHBOURS (UP TO 2 PER TARGET ADDRESS)
    # ============================================================================
    
    # Sort by distance to get closest neighbours first
    merged_df_sorted = merged_df.sort_values(['Reference_Clean', 'Distance M'])
    
    # Create neighbour columns
    def get_neighbours(group):
        """
        For each target address, get up to 2 closest neighbours
        """
        # Get non-null neighbours
        neighbours = group[group['Neighbour Address'].notna()].copy()
        
        if len(neighbours) == 0:
            return pd.Series({
                'Neighbour 1 Address': None,
                'Neighbour 1 Coordinates': None,
                'Neighbour 2 Address': None,
                'Neighbour 2 Coordinates': None,
                'Total Neighbours': 0
            })
        
        # Get first neighbour
        first = neighbours.iloc[0]
        
        # Get second neighbour if exists
        if len(neighbours) >= 2:
            second = neighbours.iloc[1]
        else:
            second = None
        
        # Format coordinates
        def format_coords(row):
            if row is None:
                return None
            lat = row.get('Neighbour Latitude')
            lon = row.get('Neighbour Longitude')
            if pd.isna(lat) or pd.isna(lon):
                return None
            return f"{lat}, {lon}"
        
        return pd.Series({
            'Neighbour 1 Address': first.get('Neighbour Address'),
            'Neighbour 1 Coordinates': format_coords(first),
            'Neighbour 2 Address': second.get('Neighbour Address') if second is not None else None,
            'Neighbour 2 Coordinates': format_coords(second) if second is not None else None,
            'Total Neighbours': len(neighbours)
        })
    
    # Apply neighbour extraction
    neighbours_df = merged_df_sorted.groupby('Reference_Clean').apply(get_neighbours).reset_index()
    
    # Merge back with original data (keeping only unique references)
    df1_unique = df1_clean.drop_duplicates('Reference_Clean')
    
    final_df = pd.merge(
        df1_unique,
        neighbours_df,
        on='Reference_Clean',
        how='left'
    )
    
    # ============================================================================
    # STEP 4: SELECT FINAL COLUMNS
    # ============================================================================
    
    # Define final columns with spaces
    final_columns = [
        'Websites',
        'Url',
        'Reference',
        'Address',
        'Proposal',
        'Application Type',
        'Application Form PDF Available',
        'Applicant = Agent',
        'Agent Name',
        'Agent Address',
        'Applicant Name',
        'Applicant Address',
        'Pdf Link',
        'Neighbour 1 Address',
        'Neighbour 1 Coordinates',
        'Neighbour 2 Address',
        'Neighbour 2 Coordinates',
        'Total Neighbours'
    ]
    
    # Ensure all columns exist
    existing_columns = [col for col in final_columns if col in final_df.columns]
    final_output = final_df[existing_columns]
    
    # ============================================================================
    # STEP 5: ADD LETTER NAME (Based on Client's IF/THEN Logic)
    # ============================================================================
    
    def generate_letter_name_and_salutation(row):
        """
        Generate Letter Name and Salutation based on client's IF/THEN logic
        Client Requirements:
        1. "Mr P Steele" -> "Dear Steele" (Use surname only)
        2. Two applicants (e.g., "Mr. Chris Banks & Mrs. Claire Tyler") -> Use full name as-is
        3. Joint applicants with "&" or "AND" -> Use full name as-is
        4. Surname only -> Add "Household"
        5. Initial as first name -> Use surname + "Household"
        6. Normal case -> Use first name only
        """
        applicant_name = row.get('Applicant Name', '')
        
        # Check for blank applicant name
        if pd.isna(applicant_name) or str(applicant_name).strip() == '':
            return pd.Series({
                'Letter Name': 'No Applicant Name',
                'Salutation': 'Dear [No Applicant Name]'
            })
        
        name_str = str(applicant_name).strip()
        name_parts = name_str.split()
        
        # Check for business exclusions (should have been filtered out, but double-check)
        business_keywords = ['Ltd', 'Limited', 'PLC', 'LLP', 'CIC', 'Group', 'Holdings', 
                           'Partnership', 'Consulting', 'Services', 'Development', 
                           'School', 'Council', 'Church', 'C/O Agent', 'c/o Agent']
        
        for keyword in business_keywords:
            if keyword.lower() in name_str.lower():
                return pd.Series({
                    'Letter Name': 'Business Excluded',
                    'Salutation': 'Dear [Business Name - Excluded]'
                })
        
        # ====================================================================
        # CLIENT'S IF/THEN LOGIC FOR LETTER NAME AND SALUTATION
        # ====================================================================
        
        # IF: Two applicant names listed (contains & or AND)
        # Example: "Mr. Chris Banks & Mrs. Claire Tyler"
        # Use full name as-is: "Dear Mr. Chris Banks & Mrs. Claire Tyler"
        if '&' in name_str or ' AND ' in name_str.upper():
            return pd.Series({
                'Letter Name': f'Letter Two - Joint: {name_str}',
                'Salutation': f'Dear {name_str}'
            })
        
        # IF: Only one name listed (surname only, e.g., "Smith")
        # Add "Household": "Dear Smith Household"
        if len(name_parts) == 1:
            return pd.Series({
                'Letter Name': f'Letter Two - Household: {name_str} Household',
                'Salutation': f'Dear {name_str} Household'
            })
        
        # IF: Only an initial used as first name (e.g., "R Smith" or "Mr P Steele")
        # Use surname + "Household": "Dear Steele Household"
        first_part = name_parts[0]
        
        # Check if first part is an initial (single letter or letter with period)
        is_initial = len(first_part) == 1 or (len(first_part) == 2 and first_part[1] == '.')
        
        # Also check for titles like "Mr", "Mrs", "Ms", "Dr" etc.
        titles = ['mr', 'mrs', 'ms', 'dr', 'prof', 'rev', 'sir', 'lord', 'lady']
        is_title = first_part.lower().replace('.', '') in titles
        
        if is_initial or is_title:
            # Find the surname (last word)
            surname = name_parts[-1] if len(name_parts) > 1 else name_str
            
            # Clean surname (remove any punctuation)
            surname = surname.strip('.,')
            
            return pd.Series({
                'Letter Name': f'Letter Two - Household: {surname} Household',
                'Salutation': f'Dear {surname} Household'
            })
        
        # ELSE: Normal case - use first name only
        # Example: "Mr Chris Banks" -> "Dear Chris"
        # But client said "Mr P Steele" -> "Dear Steele" (handled above)
        # For normal names with first name, use first name only
        first_name = name_parts[0]
        
        # Remove titles from first name if present
        for title in titles:
            if first_name.lower().replace('.', '') == title:
                # If first word is a title, use second word as first name
                if len(name_parts) > 1:
                    first_name = name_parts[1]
                else:
                    first_name = name_str
                break
        
        # If first name still has a title (e.g., "Mr."), clean it
        first_name = first_name.strip('.,')
        
        return pd.Series({
            'Letter Name': f'Letter Two - Individual: {first_name}',
            'Salutation': f'Dear {first_name}'
        })
    
    # Apply the function to create both columns
    final_output[['Letter Name', 'Salutation']] = final_output.apply(
        generate_letter_name_and_salutation, axis=1
    )
    
    # ============================================================================
    # STEP 6: APPLY BUSINESS FILTERING (Applicant Name Only)
    # ============================================================================
    
    def is_business_applicant(name):
        if pd.isna(name) or str(name).strip() == '':
            return True
        name_str = str(name).lower().strip()
        
        # Check for C/O Agent
        if 'c/o' in name_str or 'c/' in name_str:
            return True
        if 'agent' in name_str:
            return True
        
        # Business keywords
        business_keywords = [
            'ltd', 'limited', 'plc', 'llp', 'cic',
            'group', 'holdings', 'partnership',
            'consulting', 'services', 'development',
            'school', 'council', 'church',
            'estates', 'properties', 'company',
            'corporation', 'incorporated',
            'associates', 'partners',
            'trading', 'enterprises',
            'capital', 'investments', 'management',
            'solutions', 'systems', 'technologies',
            'designs', 'constructions', 'builders',
            'contractors', 'developers'
        ]
        
        for keyword in business_keywords:
            if keyword in name_str:
                return True
        return False
    
    # Identify business applicants
    final_output['Is_Business'] = final_output['Applicant Name'].apply(is_business_applicant)
    
    # Split into cleaned and removed
    cleaned_df = final_output[~final_output['Is_Business']].copy()
    removed_df = final_output[final_output['Is_Business']].copy()
    
    # Add removal reason
    def get_removal_reason(name):
        if pd.isna(name) or str(name).strip() == '':
            return 'Blank Applicant Name'
        name_str = str(name).lower()
        
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
    
    removed_df['Removal Reason'] = removed_df['Applicant Name'].apply(get_removal_reason)
    
    # Remove the temporary column
    cleaned_df = cleaned_df.drop('Is_Business', axis=1)
    removed_df = removed_df.drop('Is_Business', axis=1)
    
    # ============================================================================
    # STEP 7: SAVE OUTPUT FILES
    # ============================================================================
    
    # Save cleaned data
    cleaned_df.to_csv('final_merged_data.csv', index=False)
    print(f"\n✓ Cleaned merged data saved to: final_merged_data.csv")
    
    # Save removed business records
    removed_df.to_csv('removed_business_applicants.csv', index=False)
    print(f"✓ Removed business records saved to: removed_business_applicants.csv")
    
    # ============================================================================
    # STEP 8: GENERATE SUMMARY
    # ============================================================================
    
    summary = {
        'Metric': [
            'Total Records in df1',
            'Total Records in df2',
            'Merged Records',
            'Records with 2 Neighbours',
            'Records with 1 Neighbour',
            'Records with 0 Neighbours',
            'Business Records Removed',
            'Final Clean Records'
        ],
        'Value': [
            len(df1),
            len(df2),
            len(final_output),
            cleaned_df['Neighbour 2 Address'].notna().sum(),
            cleaned_df['Neighbour 1 Address'].notna().sum() - cleaned_df['Neighbour 2 Address'].notna().sum(),
            cleaned_df['Neighbour 1 Address'].isna().sum(),
            len(removed_df),
            len(cleaned_df)
        ]
    }
    
    summary_df = pd.DataFrame(summary)
    
    print("\n" + "="*70)
    print("SUMMARY REPORT")
    print("="*70)
    print(summary_df.to_string(index=False))
    
    # ============================================================================
    # STEP 9: PREVIEW DATA
    # ============================================================================
    
    print("\n" + "="*70)
    print("PREVIEW OF FINAL DATA (First 10 Records)")
    print("="*70)
    
    preview_cols = [
        'Reference', 'Address', 'Applicant Name', 
        'Neighbour 1 Address', 'Neighbour 1 Coordinates',
        'Neighbour 2 Address', 'Neighbour 2 Coordinates',
        'Total Neighbours', 'Letter Name', 'Salutation'
    ]
    preview_cols = [col for col in preview_cols if col in cleaned_df.columns]
    
    if len(cleaned_df) > 0:
        print(cleaned_df[preview_cols].head(10).to_string(index=False))
    else:
        print("No records after filtering!")
    
    # ============================================================================
    # STEP 10: LETTER NAME DISTRIBUTION
    # ============================================================================
    
    print("\n" + "="*70)
    print("LETTER NAME DISTRIBUTION")
    print("="*70)
    if len(cleaned_df) > 0:
        print(cleaned_df['Letter Name'].value_counts().to_string())
    
    print("\n" + "="*70)
    print("SALUTATION EXAMPLES")
    print("="*70)
    if len(cleaned_df) > 0:
        examples = cleaned_df[['Applicant Name', 'Salutation']].head(15)
        print(examples.to_string(index=False))
    
    return cleaned_df, removed_df, summary_df

# ============================================================================
# EXECUTE THE SCRIPT
# ============================================================================

# Load your dataframes
df1 = pd.read_csv(r'd:\Projects\U095\Telephone-Scrapping\applications_cleaned.csv', encoding='utf-8')  # Your df1 file
df2 = pd.read_excel(r'd:\Projects\U095\Telephone-Scrapping\Output_Neighbour\neighbours_20260818_030307.xlsx')     # Your df2 file with neighbour data

cleaned_data, removed_data, summary = merge_with_neighbours(df1, df2)

# ============================================================================
# ADDITIONAL: CHECK FOR DUPLICATES
# ============================================================================

print("\n" + "="*70)
print("DUPLICATE CHECK")
print("="*70)

# Check for duplicate References
duplicates = cleaned_data[cleaned_data.duplicated(['Reference'], keep=False)]
if len(duplicates) > 0:
    print(f"Found {len(duplicates)} duplicate references:")
    print(duplicates[['Reference', 'Address', 'Applicant Name']].head())
else:
    print("No duplicate references found.")