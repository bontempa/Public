# Public
import pandas as pd
import re
import difflib

def normalize_name(name):
    """
    Normalize company names by:
    - Converting to lowercase
    - Removing punctuation and special characters
    - Removing common corporate suffixes
    - Removing extra whitespace
    """
    name = str(name).lower()
    name = re.sub(r'[^\w\s]', '', name)  # Remove punctuation
    name = re.sub(r'\b(inc|ltd|llc|corp|co|company|international|intl)\b', '', name)
    name = re.sub(r'\s+', ' ', name)  # Replace multiple spaces with a single space
    return name.strip()

def match_company_names(manual_names, official_names, cutoff=0.6):
    """
    Match company names using difflib.

    Parameters:
    - manual_names: List of names to search for.
    - official_names: Reference list of official company names.
    - cutoff: Minimum similarity ratio (0 to 1) for a match.

    Returns:
    - DataFrame with original names, matched names, and similarity scores.
    """
    # Normalize the lists of names
    normalized_manual = [normalize_name(name) for name in manual_names]
    normalized_official = [normalize_name(name) for name in official_names]

    results = []
    for original_name, normalized_name in zip(manual_names, normalized_manual):
        # Perform fuzzy matching using difflib
        matches = difflib.get_close_matches(normalized_name, normalized_official, n=1, cutoff=cutoff)
        if matches:
            match = matches[0]
            # Find the original official name corresponding to the normalized match
            matched_official_name = official_names[normalized_official.index(match)]
            # Calculate a similarity score
            score = difflib.SequenceMatcher(None, normalized_name, match).ratio() * 100
            results.append({
                'Original Name': original_name,
                'Matched Name': matched_official_name,
                'Score': round(score, 2)
            })
        else:
            results.append({
                'Original Name': original_name,
                'Matched Name': None,
                'Score': None
            })
    return pd.DataFrame(results)

# Read data from Excel
# Replace 'company_names.xlsx' with the path to your Excel file
excel_file = 'company_names.xlsx'

# Read the 'investors' sheet column A
investors_df = pd.read_excel(excel_file, sheet_name='investors', usecols='A')
manual_names = investors_df.iloc[:, 0].dropna().tolist()

# Read the 'ref' sheet column A
ref_df = pd.read_excel(excel_file, sheet_name='ref', usecols='A')
official_names = ref_df.iloc[:, 0].dropna().tolist()

# Perform matching
df_matches = match_company_names(manual_names, official_names)

# Output the results
print(df_matches)

# Optionally, save the results to a new Excel file
df_matches.to_excel('matched_companies.xlsx', index=False)
