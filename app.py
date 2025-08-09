import streamlit as st
import pandas as pd
import re
import io
from typing import Dict, List, Tuple
from difflib import SequenceMatcher
import xlsxwriter

# Page configuration
st.set_page_config(
    page_title="Job Status Job Description Merging Tool",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded"
)

def normalize_key(series):
    """
    Normalize keys according to the specified rules:
    - Trim whitespace
    - Convert to uppercase
    - Collapse multiple spaces to single space
    - Remove leading/trailing punctuation
    - Treat empty strings as nulls
    """
    def normalize_single_value(value):
        if pd.isna(value) or value == '':
            return None
        
        # Convert to string and strip whitespace
        normalized = str(value).strip()
        
        # Convert to uppercase
        normalized = normalized.upper()
        
        # Collapse multiple spaces to single space
        normalized = re.sub(r'\s+', ' ', normalized)
        
        # Remove leading/trailing punctuation
        normalized = re.sub(r'^[^\w\s]+|[^\w\s]+$', '', normalized)
        
        # Final trim
        normalized = normalized.strip()
        
        # Return None for empty strings
        return normalized if normalized else None
    
    return series.apply(normalize_single_value)



def extract_vessel_name(filename: str) -> str:
    """Extract vessel name from filename by removing file extension and common suffixes."""
    if not filename:
        return "Unknown Vessel"
    
    # Remove file extension
    name = filename.rsplit('.', 1)[0]
    
    # Remove common suffixes like dates, timestamps, etc.
    # Look for patterns like " 08082025", "_1754713537058", etc.
    name = re.sub(r'[\s_]\d{8,}[\s_]*.*$', '', name)
    name = re.sub(r'[\s_]\d{2}\d{2}\d{4}[\s_]*.*$', '', name)
    
    # Clean up any trailing underscores or spaces
    name = re.sub(r'[\s_]+$', '', name)
    
    return name.strip() if name.strip() else "Unknown Vessel"

def create_excel_export(enriched_data: pd.DataFrame, vessel_name: str) -> bytes:
    """Create a formatted Excel file with the enriched data."""
    
    # Define the exact column order as requested
    column_order = [
        'Vessel',
        'Function',
        'Machinery Location', 
        'Sub Component Location',
        'Job Code',
        'Title',
        'Description',
        'Frequency',
        'Performing Rank',
        'Verifying Rank', 
        'Maker',
        'Model',
        'Calculated Due Date',
        'Due Date',
        'Next Due',
        'Completion Date',
        'Job Status',
        'Job Action',
        'Remaining Running Hours',
        'Machinery Running Hours',
        'Last Done Running Hours',
        'Last Done Date',
        'CMS Code',
        'Job Source',
        'E-Form',
        'Attachment Indicator'
    ]
    
    # Create a copy of the data for export
    export_data = enriched_data.copy()
    
    # Add vessel name column
    export_data['Vessel'] = vessel_name
    
    # Create the final export DataFrame with proper mapping
    final_export = pd.DataFrame()
    
    # Map each target column to the best matching source column
    for target_col in column_order:
        if target_col == 'Vessel':
            final_export[target_col] = [vessel_name] * len(export_data)
        elif target_col == 'Function':
            # Find function column or set empty
            function_col = None
            for col in export_data.columns:
                if col.lower() == 'function':
                    function_col = col
                    break
            final_export[target_col] = export_data[function_col] if function_col else ''
        elif target_col == 'Machinery Location':
            # Find machinery location column
            machinery_col = None
            for col in export_data.columns:
                if 'machinery' in col.lower() and ('location' in col.lower() or 'loc' in col.lower()):
                    machinery_col = col
                    break
            final_export[target_col] = export_data[machinery_col] if machinery_col else ''
        elif target_col == 'Job Code':
            # Find job code column
            job_code_col = None
            for col in export_data.columns:
                if 'job' in col.lower() and 'code' in col.lower():
                    job_code_col = col
                    break
            final_export[target_col] = export_data[job_code_col] if job_code_col else ''
        elif target_col == 'Description':
            # Find description column
            desc_col = None
            for col in export_data.columns:
                if col.lower() == 'description':
                    desc_col = col
                    break
            final_export[target_col] = export_data[desc_col] if desc_col else ''
        elif target_col == 'Maker':
            # Find maker column
            maker_col = None
            for col in export_data.columns:
                if col.lower() == 'maker':
                    maker_col = col
                    break
            final_export[target_col] = export_data[maker_col] if maker_col else ''
        elif target_col == 'Model':
            # Find model column
            model_col = None
            for col in export_data.columns:
                if col.lower() == 'model':
                    model_col = col
                    break
            final_export[target_col] = export_data[model_col] if model_col else ''
        else:
            # For other columns, try to find a matching column or set empty
            found_col = None
            target_lower = target_col.lower()
            for col in export_data.columns:
                col_lower = col.lower().replace('_', ' ').replace('-', ' ')
                if target_lower == col_lower:
                    found_col = col
                    break
                elif target_lower.replace(' ', '') == col_lower.replace(' ', ''):
                    found_col = col
                    break
            final_export[target_col] = export_data[found_col] if found_col else ''
    
    # Clean data for Excel export - replace NaN with empty strings
    final_export = final_export.fillna('')
    
    # Create Excel file in memory
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Write data to Excel
        sheet_name = f"{vessel_name} Job List"
        final_export.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2)
        
        # Get workbook and worksheet objects
        workbook = writer.book
        workbook.nan_inf_to_errors = True  # Set NaN handling option on workbook
        worksheet = writer.sheets[sheet_name]
        
        # Define formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 16,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#366092',
            'font_color': 'white',
            'border': 1
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'font_size': 11,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#D9E2F3',
            'border': 1,
            'text_wrap': True
        })
        
        cell_format = workbook.add_format({
            'align': 'left',
            'valign': 'vcenter',
            'border': 1,
            'text_wrap': True
        })
        
        # Write title
        worksheet.merge_range('A1:Z1', f"{vessel_name} Job List", title_format)
        
        # Format header row (row 1, 0-indexed) - overwrite pandas headers
        for col_num in range(len(column_order)):
            worksheet.write(1, col_num, column_order[col_num], header_format)
        
        # Apply border formatting to all data cells
        last_row = len(final_export) + 2
        for row_num in range(2, last_row + 1):  # Start from row 3 (0-indexed row 2)
            for col_num in range(len(column_order)):
                # Get existing cell value and rewrite with formatting
                cell_value = final_export.iloc[row_num - 2, col_num] if row_num - 2 < len(final_export) else ''
                worksheet.write(row_num, col_num, cell_value, cell_format)
        
        # Set column widths
        column_widths = {
            'A': 15,  # Vessel
            'B': 15,  # Function
            'C': 20,  # Machinery Location
            'D': 18,  # Sub Component Location
            'E': 12,  # Job Code
            'F': 25,  # Title
            'G': 30,  # Description
            'H': 12,  # Frequency
            'I': 15,  # Performing Rank
            'J': 15,  # Verifying Rank
            'K': 15,  # Maker
            'L': 15,  # Model
            'M': 15,  # Calculated Due Date
            'N': 12,  # Due Date
            'O': 12,  # Next Due
            'P': 15,  # Completion Date
            'Q': 12,  # Job Status
            'R': 12,  # Job Action
            'S': 18,  # Remaining Running Hours
            'T': 20,  # Machinery Running Hours
            'U': 20,  # Last Done Running Hours
            'V': 15,  # Last Done Date
            'W': 12,  # CMS Code
            'X': 12,  # Job Source
            'Y': 10,  # E-Form
            'Z': 18   # Attachment Indicator
        }
        
        for col_letter, width in column_widths.items():
            worksheet.set_column(f'{col_letter}:{col_letter}', width)
        
        # Set row heights
        worksheet.set_row(0, 25)  # Title row
        worksheet.set_row(1, 20)  # Header row
    
    output.seek(0)
    return output.getvalue()

def enrich_job_status(job_status_df: pd.DataFrame, job_list_df: pd.DataFrame, cfg: Dict) -> Dict:
    """
    Main function to enrich job status data with job list information.
    Returns a dictionary containing enriched data and diagnostics.
    """
    try:
        # Extract configuration
        js_job_code_col = cfg['js_job_code']
        js_machinery_col = cfg['js_machinery']
        jl_job_code_col = cfg['jl_job_code']
        jl_machinery_col = cfg['jl_machinery']
        jl_description_col = cfg['jl_description']
        jl_maker_col = cfg['jl_maker']
        jl_model_col = cfg['jl_model']
        keep_only_matched = cfg.get('keep_only_matched', False)
        strict_join = cfg.get('strict_join', True)
        
        # Create working copies
        js_work = job_status_df.copy()
        jl_work = job_list_df.copy()
        
        # Normalize keys
        js_work['_norm_job_code'] = normalize_key(js_work[js_job_code_col])
        js_work['_norm_machinery'] = normalize_key(js_work[js_machinery_col])
        jl_work['_norm_job_code'] = normalize_key(jl_work[jl_job_code_col])
        jl_work['_norm_machinery'] = normalize_key(jl_work[jl_machinery_col])
        
        # Create composite keys for joining
        js_work['_join_key'] = js_work['_norm_job_code'].astype(str) + '||' + js_work['_norm_machinery'].astype(str)
        jl_work['_join_key'] = jl_work['_norm_job_code'].astype(str) + '||' + jl_work['_norm_machinery'].astype(str)
        
        # Detect duplicates in job list
        jl_duplicates = jl_work[jl_work.duplicated(subset=['_join_key'], keep=False)]
        duplicate_stats = jl_duplicates.groupby('_join_key').size().reset_index()
        duplicate_stats.columns = ['_join_key', 'count']
        
        # Remove duplicates from job list (keep first occurrence)
        jl_work_dedup = jl_work.drop_duplicates(subset=['_join_key'], keep='first')
        
        # Prepare columns to merge
        merge_columns = ['_join_key', jl_description_col, jl_maker_col, jl_model_col]
        jl_merge = jl_work_dedup[merge_columns].copy()
        
        # Rename columns to avoid conflicts
        column_mapping = {
            jl_description_col: 'Description',
            jl_maker_col: 'Maker',
            jl_model_col: 'Model'
        }
        jl_merge = jl_merge.rename(columns=column_mapping)
        
        # Perform left join
        enriched = js_work.merge(jl_merge, on='_join_key', how='left')
        
        # Handle fallback join if enabled
        join_used_fallback = pd.Series([False] * len(enriched), index=enriched.index)
        
        if not strict_join:
            # Find rows that didn't match and try job code only fallback
            unmatched_mask = enriched['Description'].isna()
            unmatched_rows = enriched[unmatched_mask]
            
            if len(unmatched_rows) > 0:
                # Prepare job list for job code only matching
                jl_fallback = jl_work_dedup.groupby('_norm_job_code').first().reset_index()
                jl_fallback_merge = jl_fallback[['_norm_job_code', jl_description_col, jl_maker_col, jl_model_col]].copy()
                fallback_column_mapping = {
                    jl_description_col: 'Description_fallback',
                    jl_maker_col: 'Maker_fallback',
                    jl_model_col: 'Model_fallback'
                }
                jl_fallback_merge = jl_fallback_merge.rename(columns=fallback_column_mapping)
                
                # Merge fallback data
                enriched = enriched.merge(jl_fallback_merge, on='_norm_job_code', how='left')
                
                # Apply fallback values where original match failed
                fallback_mask = enriched['Description'].isna() & enriched['Description_fallback'].notna()
                enriched.loc[fallback_mask, 'Description'] = enriched.loc[fallback_mask, 'Description_fallback']
                enriched.loc[fallback_mask, 'Maker'] = enriched.loc[fallback_mask, 'Maker_fallback']
                enriched.loc[fallback_mask, 'Model'] = enriched.loc[fallback_mask, 'Model_fallback']
                join_used_fallback.loc[fallback_mask] = True
                
                # Drop fallback columns
                enriched = enriched.drop(columns=['Description_fallback', 'Maker_fallback', 'Model_fallback'])
            
            # Add fallback indicator column
            enriched['JoinUsedFallback'] = join_used_fallback
        
        # Clean up working columns
        enriched = enriched.drop(columns=['_norm_job_code', '_norm_machinery', '_join_key'])
        
        # Calculate initial statistics before auto-matching
        total_js_rows = len(job_status_df)
        initial_matched_rows = len(enriched[enriched['Description'].notna()])
        
        # Filter to matched rows only if requested
        if keep_only_matched:
            enriched = enriched[enriched['Description'].notna()]
        
        # Get unmatched row details
        unmatched_details = enriched[enriched['Description'].isna()][[js_job_code_col, js_machinery_col]].copy()
        
        # Generate fuzzy match suggestions for unmatched rows
        fuzzy_suggestions = pd.DataFrame()
        auto_matched_count = 0
        
        if len(unmatched_details) > 0:
            # Get unique unmatched machinery locations
            unmatched_machinery = unmatched_details[js_machinery_col].dropna().unique().tolist()
            # Get all available machinery locations from job list
            available_machinery = jl_work[jl_machinery_col].dropna().unique().tolist()
            
            # Find fuzzy matches
            suggestions = find_fuzzy_matches(unmatched_machinery, available_machinery, threshold=0.6)
            
            # Format suggestions into a DataFrame
            if suggestions:
                fuzzy_suggestions = pd.DataFrame(suggestions, columns=[
                    'Unmatched_Machinery', 'Suggested_Match', 'Similarity_Score'
                ])
                fuzzy_suggestions['Similarity_Score'] = fuzzy_suggestions['Similarity_Score'].round(3)
                
                # Additional validation check for suggestions
                validated_suggestions = []
                for _, row in fuzzy_suggestions.iterrows():
                    unmatched = row['Unmatched_Machinery']
                    suggested = row['Suggested_Match']
                    score = row['Similarity_Score']
                    
                    # Validate that the suggested match actually exists in Job List with matching Job Codes
                    unmatched_job_codes = unmatched_details[unmatched_details[js_machinery_col] == unmatched][js_job_code_col].unique()
                    suggested_entries = jl_work[jl_work[jl_machinery_col] == suggested]
                    
                    if len(suggested_entries) > 0:
                        # Check if there are compatible job codes
                        suggested_job_codes = suggested_entries[jl_job_code_col].unique()
                        has_compatible_codes = any(code in suggested_job_codes for code in unmatched_job_codes)
                        
                        if has_compatible_codes:
                            validated_suggestions.append({
                                'Unmatched_Machinery': unmatched,
                                'Suggested_Match': suggested,
                                'Similarity_Score': score,
                                'Compatible_Job_Codes': len(set(unmatched_job_codes) & set(suggested_job_codes)),
                                'Validation_Status': 'Valid'
                            })
                        else:
                            validated_suggestions.append({
                                'Unmatched_Machinery': unmatched,
                                'Suggested_Match': suggested,
                                'Similarity_Score': score,
                                'Compatible_Job_Codes': 0,
                                'Validation_Status': 'No Compatible Job Codes'
                            })
                
                if validated_suggestions:
                    fuzzy_suggestions = pd.DataFrame(validated_suggestions)
                    fuzzy_suggestions['Similarity_Score'] = fuzzy_suggestions['Similarity_Score'].round(3)
                    
                    # Apply high confidence matches (≥85% similarity) with validation
                    high_confidence_matches = fuzzy_suggestions[
                        (fuzzy_suggestions['Similarity_Score'] >= 0.85) & 
                        (fuzzy_suggestions['Validation_Status'] == 'Valid')
                    ]
                    
                    if len(high_confidence_matches) > 0:
                        # Create a mapping dictionary for high confidence matches
                        auto_match_dict = {}
                        for _, row in high_confidence_matches.iterrows():
                            # Normalize both for consistent matching
                            unmatched_norm = normalize_key(pd.Series([row['Unmatched_Machinery']])).iloc[0]
                            suggested_norm = normalize_key(pd.Series([row['Suggested_Match']])).iloc[0]
                            auto_match_dict[unmatched_norm] = suggested_norm
                        
                        # Re-create enriched data with updated join keys for auto-matching
                        js_auto = js_work.copy()
                        
                        # Update join keys for high confidence matches
                        for unmatched_norm, suggested_norm in auto_match_dict.items():
                            # Find rows in Job Status that have the unmatched machinery (normalized)
                            mask = js_auto['_norm_machinery'] == unmatched_norm
                            if mask.any():
                                # Update the normalized machinery to match the suggested one
                                js_auto.loc[mask, '_norm_machinery'] = suggested_norm
                                # Update the join key accordingly
                                js_auto.loc[mask, '_join_key'] = js_auto.loc[mask, '_norm_job_code'].astype(str) + '||' + suggested_norm
                        
                        # Perform the merge again with updated join keys
                        enriched = js_auto.merge(jl_merge, on='_join_key', how='left')
                        
                        # Count how many were auto-matched
                        auto_matched_count = len(enriched[enriched['Description'].notna()]) - initial_matched_rows
                        
                        # Remove auto-matched entries from fuzzy suggestions for display
                        if auto_matched_count > 0:
                            # Filter out auto-matched entries from suggestions
                            auto_matched_original = high_confidence_matches['Unmatched_Machinery'].tolist()
                            fuzzy_suggestions = fuzzy_suggestions[~fuzzy_suggestions['Unmatched_Machinery'].isin(auto_matched_original)]
                    
                    # Remove validation columns for cleaner display but keep the info for debugging
                    display_suggestions = fuzzy_suggestions[['Unmatched_Machinery', 'Suggested_Match', 'Similarity_Score']].copy()
                    fuzzy_suggestions = display_suggestions
        
        # Additional matching step: Job Code only for remaining unmatched records
        job_code_only_matched_count = 0
        if len(enriched[enriched['Description'].isna()]) > 0:
            # Get currently unmatched rows
            still_unmatched = enriched[enriched['Description'].isna()].copy()
            
            # Prepare Job List data for Job Code only matching
            jl_job_code_only = jl_work[[jl_job_code_col, jl_description_col, jl_maker_col, jl_model_col, '_norm_job_code']].copy()
            jl_job_code_only = jl_job_code_only.drop_duplicates(subset=['_norm_job_code'], keep='first')
            
            # Rename columns for merge
            jl_job_code_only_merge = jl_job_code_only[['_norm_job_code', jl_description_col, jl_maker_col, jl_model_col]].copy()
            column_mapping_job_code = {
                jl_description_col: 'Description_JobCode',
                jl_maker_col: 'Maker_JobCode', 
                jl_model_col: 'Model_JobCode'
            }
            jl_job_code_only_merge = jl_job_code_only_merge.rename(columns=column_mapping_job_code)
            
            # Merge unmatched rows with Job List using only Job Code
            unmatched_with_job_code = still_unmatched.merge(jl_job_code_only_merge, on='_norm_job_code', how='left')
            
            # Find which rows in the unmatched data now have job code matches
            job_code_found_mask = unmatched_with_job_code['Description_JobCode'].notna()
            
            if job_code_found_mask.any():
                # Count matches found by Job Code only
                job_code_only_matched_count = job_code_found_mask.sum()
                
                # Get the indices of rows that were matched
                matched_indices = still_unmatched.index[job_code_found_mask]
                
                # Update the enriched data with Job Code only matches using proper indexing
                enriched.loc[matched_indices, 'Description'] = unmatched_with_job_code.loc[job_code_found_mask, 'Description_JobCode'].values
                enriched.loc[matched_indices, 'Maker'] = unmatched_with_job_code.loc[job_code_found_mask, 'Maker_JobCode'].values
                enriched.loc[matched_indices, 'Model'] = unmatched_with_job_code.loc[job_code_found_mask, 'Model_JobCode'].values
                
                # Mark these as using fallback (Job Code only) matching
                if 'JoinUsedFallback' in enriched.columns:
                    enriched.loc[matched_indices, 'JoinUsedFallback'] = True
                else:
                    enriched['JoinUsedFallback'] = False
                    enriched.loc[matched_indices, 'JoinUsedFallback'] = True
        
        # Calculate final statistics after all matching attempts
        final_matched_rows = len(enriched[enriched['Description'].notna()])
        final_unmatched_rows = total_js_rows - final_matched_rows
        
        # Get unmatched row details after all matching attempts
        unmatched_details = enriched[enriched['Description'].isna()][[js_job_code_col, js_machinery_col]].copy()
        
        # Prepare duplicate details
        if len(duplicate_stats) > 0:
            duplicate_details = []
            for _, row in duplicate_stats.iterrows():
                key_parts = str(row['_join_key']).split('||')
                if len(key_parts) == 2:
                    duplicate_details.append({
                        'Job Code': key_parts[0],
                        'Machinery': key_parts[1],
                        'Duplicate Count': int(row['count'])
                    })
            duplicate_details_df = pd.DataFrame(duplicate_details)
        else:
            duplicate_details_df = pd.DataFrame()
        
        return {
            'enriched_data': enriched,
            'total_js_rows': total_js_rows,
            'total_jl_rows': len(job_list_df),
            'matched_rows': final_matched_rows,
            'unmatched_rows': final_unmatched_rows,
            'auto_matched_rows': auto_matched_count,
            'job_code_only_matched_rows': job_code_only_matched_count,
            'unmatched_details': unmatched_details,
            'duplicate_details': duplicate_details_df,
            'fuzzy_suggestions': fuzzy_suggestions,
            'success': True,
            'error': None
        }
        
    except Exception as e:
        return {
            'success': False,
            'error': str(e)
        }

def similarity_score(a: str, b: str) -> float:
    """Calculate similarity score between two strings with enhanced matching for maritime equipment patterns."""
    # Basic similarity
    basic_score = SequenceMatcher(None, a, b).ratio()
    
    # Enhanced scoring for common abbreviation patterns
    a_clean = a.upper().strip()
    b_clean = b.upper().strip()
    
    # Remove common trailing patterns for better matching
    def normalize_for_comparison(text):
        # Remove trailing spaces and common suffixes
        text = text.strip()
        
        # Handle specific Liferaft patterns first
        import re
        # LiferaftStarboard1/2 -> Liferaft
        if 'LIFERAFT' in text:
            # Extract base equipment name for Liferaft variations
            if re.search(r'LIFERAFTSTARBOARD\d+', text):
                return 'LIFERAFT'
            elif re.search(r'LIFERAFTS\d+', text):
                return 'LIFERAFT'
            elif re.search(r'LIFERAFT#\d+', text):
                return 'LIFERAFT'
        
        # Handle trailing single characters that might be abbreviations
        if len(text) > 1 and text[-1] in ['P', 'S', 'T'] and text[-2] == ' ':
            text = text[:-1].strip()
        # Handle number suffixes like #1, #2, etc.
        text = re.sub(r'#\d+$', '', text).strip()
        # Handle trailing "Port" vs "P" abbreviations
        if text.endswith('PORT'):
            text = text[:-4].strip()
        elif text.endswith('P ') or text.endswith('P'):
            text = text[:-1].strip()
        # Handle "Starboard" variations
        if text.endswith('STARBOARD'):
            text = text[:-9].strip()
        elif text.endswith('STBD') or text.endswith('STB'):
            text = text[:-4].strip() if text.endswith('STBD') else text[:-3].strip()
        
        # Remove trailing numbers for general equipment matching
        text = re.sub(r'\d+$', '', text).strip()
        
        return text
    
    # Normalize both strings for comparison
    a_norm = normalize_for_comparison(a_clean)
    b_norm = normalize_for_comparison(b_clean)
    
    # Calculate normalized similarity
    normalized_score = SequenceMatcher(None, a_norm, b_norm).ratio()
    
    # Common abbreviation mappings for maritime equipment
    abbreviations = {
        'STARBOARD': ['STBD', 'STB', 'S'],
        'PORT': ['PT', 'P'],
        'FORWARD': ['FWD'],
        'AFTER': ['AFT'],
        'SPACES': ['SPACE'],
        'LADDER': ['LDR'],
        'ENGINE': ['ENG'],
        'SYSTEM': ['SYS'],
        'EQUIPMENT': ['EQUIP', 'EQ'],
        'COMBINATION': ['COMBO'],
        'DAVIT': ['DAV'],
        'PILOT': ['PLT']
    }
    
    # Special equipment name mappings based on your suggestions
    special_mappings = {
        'LIFERAFTSTARBOARD1': ['LIFERAFTS1', 'LIFERAFT#1'],
        'LIFERAFTSTARBOARD2': ['LIFERAFTS2', 'LIFERAFT#2', 'LIFERAFTS1'],  # Cross matching as suggested
        'PILOT COMBINATION LADDERPORT': ['PILOT COMBINATION LADDERP'],
        'BUNKER DAVITSTARBOARD': ['BUNKER DAVITSTBD'],
        'CHAIN LOCKERSTARBOARD': ['CHAIN LOCKERSTBD']
    }
    
    # Check for special equipment mappings first
    special_boost = 0
    for equipment, variants in special_mappings.items():
        if (equipment in a_clean and any(variant in b_clean for variant in variants)) or \
           (equipment in b_clean and any(variant in a_clean for variant in variants)) or \
           (any(variant in a_clean for variant in variants) and any(variant in b_clean for variant in variants)):
            special_boost = 0.25  # High boost for known equipment mappings
            
    # Check if one string is an abbreviation of another
    abbrev_boost = 0
    for full_word, abbrevs in abbreviations.items():
        for abbrev in abbrevs:
            if (full_word in a_clean and abbrev in b_clean) or (full_word in b_clean and abbrev in a_clean):
                abbrev_boost = 0.15  # Boost for abbreviation matches
    
    # Check for exact substring matches (case insensitive)
    substring_boost = 0
    if a_clean in b_clean or b_clean in a_clean:
        substring_boost = 0.1
    
    # Special handling for very similar equipment names
    equipment_boost = 0
    if normalized_score >= 0.8:
        equipment_boost = 0.1
    
    # Take the best score from all methods
    final_score = max(basic_score, normalized_score + special_boost + abbrev_boost + substring_boost + equipment_boost)
    
    # Cap at 1.0
    return min(final_score, 1.0)

def find_fuzzy_matches(unmatched_keys: List[str], available_keys: List[str], threshold: float = 0.6) -> List[Tuple[str, str, float]]:
    """Find potential fuzzy matches for unmatched keys."""
    suggestions = []
    
    for unmatched in unmatched_keys:
        best_matches = []
        
        for available in available_keys:
            score = similarity_score(unmatched.upper(), available.upper())
            if score >= threshold:
                best_matches.append((available, score))
        
        # Sort by score and take top matches
        best_matches.sort(key=lambda x: x[1], reverse=True)
        
        for match, score in best_matches[:3]:  # Top 3 matches
            suggestions.append((unmatched, match, score))
    
    return suggestions

def convert_df_to_csv(df: pd.DataFrame) -> str:
    """Convert a DataFrame to CSV string for download."""
    return df.to_csv(index=False, encoding='utf-8')

def main():
    st.title("🔧 Job Status Enrichment Tool")
    st.markdown("Upload Job Status and Job List CSV files to enrich job status data with additional information.")
    
    # Initialize session state
    if 'job_status_df' not in st.session_state:
        st.session_state.job_status_df = None
    if 'job_list_df' not in st.session_state:
        st.session_state.job_list_df = None
    if 'enrichment_result' not in st.session_state:
        st.session_state.enrichment_result = None
    

    
    # File upload section
    st.header("📁 File Upload")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Job Status Export")
        job_status_file = st.file_uploader(
            "Upload Job Status CSV",
            type=['csv'],
            key="job_status_upload"
        )
        
        if job_status_file is not None:
            try:
                st.session_state.job_status_df = pd.read_csv(job_status_file)
                st.session_state.job_status_file = job_status_file  # Store file reference for vessel name extraction
                st.success(f"✅ Loaded {len(st.session_state.job_status_df)} rows")
                
                # Show preview
                st.subheader("Preview (First 10 rows)")
                st.dataframe(st.session_state.job_status_df.head(10), use_container_width=True)
                
            except Exception as e:
                st.error(f"❌ Error loading Job Status file: {str(e)}")
                st.session_state.job_status_df = None
    
    with col2:
        st.subheader("Job List Export")
        job_list_file = st.file_uploader(
            "Upload Job List CSV",
            type=['csv'],
            key="job_list_upload"
        )
        
        if job_list_file is not None:
            try:
                st.session_state.job_list_df = pd.read_csv(job_list_file)
                st.success(f"✅ Loaded {len(st.session_state.job_list_df)} rows")
                
                # Show preview
                st.subheader("Preview (First 10 rows)")
                st.dataframe(st.session_state.job_list_df.head(10), use_container_width=True)
                
            except Exception as e:
                st.error(f"❌ Error loading Job List file: {str(e)}")
                st.session_state.job_list_df = None
    
    # Column mapping section
    if st.session_state.job_status_df is not None and st.session_state.job_list_df is not None:
        st.header("🗂️ Column Mapping")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Job Status Columns")
            js_columns = list(st.session_state.job_status_df.columns)
            
            js_job_code = st.selectbox(
                "Job Code Column",
                js_columns,
                index=js_columns.index('Job Code') if 'Job Code' in js_columns else 0,
                key="js_job_code"
            )
            
            js_machinery = st.selectbox(
                "Machinery Location Column",
                js_columns,
                index=js_columns.index('Machinery Location') if 'Machinery Location' in js_columns else 0,
                key="js_machinery"
            )
        
        with col2:
            st.subheader("Job List Columns")
            jl_columns = list(st.session_state.job_list_df.columns)
            
            jl_job_code = st.selectbox(
                "Job Code Column",
                jl_columns,
                index=jl_columns.index('Job Code') if 'Job Code' in jl_columns else 0,
                key="jl_job_code"
            )
            
            jl_machinery = st.selectbox(
                "Machinery Column",
                jl_columns,
                index=jl_columns.index('Machinery') if 'Machinery' in jl_columns else 0,
                key="jl_machinery"
            )
            
            jl_description = st.selectbox(
                "Description Column",
                jl_columns,
                index=jl_columns.index('Description') if 'Description' in jl_columns else 0,
                key="jl_description"
            )
            
            jl_maker = st.selectbox(
                "Maker Column",
                jl_columns,
                index=jl_columns.index('Maker') if 'Maker' in jl_columns else 0,
                key="jl_maker"
            )
            
            jl_model = st.selectbox(
                "Model Column",
                jl_columns,
                index=jl_columns.index('Model') if 'Model' in jl_columns else 0,
                key="jl_model"
            )
        
        # Options section
        st.header("⚙️ Options")
        
        col1, col2 = st.columns(2)
        
        with col1:
            keep_only_matched = st.checkbox(
                "Keep only matched rows",
                help="Filter output to show only rows that found matches in Job List"
            )
        
        with col2:
            strict_join = st.checkbox(
                "Strict join (both keys required)",
                value=True,
                help="Uncheck to enable fallback join by Job Code only"
            )
        
        # Validation
        st.header("🔍 Validation")
        
        validation_errors = []
        
        # Check if required columns exist
        required_js_cols = [js_job_code, js_machinery]
        required_jl_cols = [jl_job_code, jl_machinery, jl_description, jl_maker, jl_model]
        
        missing_js = [col for col in required_js_cols if col not in js_columns]
        missing_jl = [col for col in required_jl_cols if col not in jl_columns]
        
        if missing_js:
            validation_errors.append(f"Missing columns in Job Status: {missing_js}")
        
        if missing_jl:
            validation_errors.append(f"Missing columns in Job List: {missing_jl}")
        
        if validation_errors:
            for error in validation_errors:
                st.error(f"❌ {error}")
        else:
            st.success("✅ All required columns found!")
        
        # Run merge button
        if st.button("🚀 Run Merge", disabled=bool(validation_errors), type="primary"):
            with st.spinner("Processing merge..."):
                config = {
                    'js_job_code': js_job_code,
                    'js_machinery': js_machinery,
                    'jl_job_code': jl_job_code,
                    'jl_machinery': jl_machinery,
                    'jl_description': jl_description,
                    'jl_maker': jl_maker,
                    'jl_model': jl_model,
                    'keep_only_matched': keep_only_matched,
                    'strict_join': strict_join
                }
                
                st.session_state.enrichment_result = enrich_job_status(
                    st.session_state.job_status_df,
                    st.session_state.job_list_df,
                    config
                )
        
        # Results section
        if st.session_state.enrichment_result is not None:
            result = st.session_state.enrichment_result
            
            if result['success']:
                st.header("📊 Results")
                
                # Success banner
                st.success("🎉 Merge completed successfully!")
                
                # Auto-matching info
                if result['auto_matched_rows'] > 0:
                    st.info(f"✨ **Smart Matching Applied!** Automatically matched {result['auto_matched_rows']} additional rows using high-confidence fuzzy matching (≥85% similarity).")
                
                # Job Code only matching info
                if result.get('job_code_only_matched_rows', 0) > 0:
                    st.info(f"🔍 **Fallback Matching Applied!** Matched {result['job_code_only_matched_rows']} additional rows using Job Code only (ignoring machinery location differences).")
                
                # Statistics
                st.subheader("📈 Statistics")
                
                col1, col2, col3, col4, col5, col6 = st.columns(6)
                
                with col1:
                    st.metric("Job Status Rows", result['total_js_rows'])
                
                with col2:
                    st.metric("Job List Rows", result['total_jl_rows'])
                
                with col3:
                    st.metric("Matched Rows", result['matched_rows'])
                
                with col4:
                    st.metric("Auto-Matched", result['auto_matched_rows'], help="Rows matched using high confidence fuzzy matching (≥85% similarity)")
                
                with col5:
                    st.metric("Job Code Only", result.get('job_code_only_matched_rows', 0), help="Rows matched using Job Code only (fallback matching)")
                
                with col6:
                    st.metric("Still Unmatched", result['unmatched_rows'])
                
                # Enriched data preview
                st.subheader("🔧 Enriched Data Preview")
                st.dataframe(result['enriched_data'].head(20), use_container_width=True)
                
                # Diagnostics
                st.subheader("🔍 Diagnostics")
                
                # Unmatched rows
                if len(result['unmatched_details']) > 0:
                    st.warning(f"⚠️ {len(result['unmatched_details'])} unmatched rows found")
                    
                    # Show fuzzy match suggestions
                    if len(result['fuzzy_suggestions']) > 0:
                        st.info("💡 **Suggested Matches Found!** Consider these potential matches:")
                        with st.expander("🔍 View Matching Suggestions", expanded=True):
                            st.markdown("These suggestions are based on text similarity analysis:")
                            
                            # Show only high confidence suggestions
                            high_confidence = result['fuzzy_suggestions'][result['fuzzy_suggestions']['Similarity_Score'] >= 0.85]
                            
                            if len(high_confidence) > 0:
                                st.success("**High Confidence Matches (≥85% similarity):**")
                                st.dataframe(high_confidence, use_container_width=True, hide_index=True)
                            
                            st.markdown("""
                            **How to use these suggestions:**
                            1. Review the suggested matches above
                            2. Update your Job List CSV file to use consistent naming
                            3. For example: `Lifeboat/Rescue BoatStarboard` → `Lifeboat/Rescue BoatStbd`
                            4. Re-upload the updated Job List file and run the merge again
                            """)
                    
                    with st.expander("View All Unmatched Rows"):
                        st.dataframe(result['unmatched_details'], use_container_width=True)
                else:
                    st.success("✅ All rows matched successfully!")
                
                # Duplicate keys
                if len(result['duplicate_details']) > 0:
                    st.warning(f"⚠️ {len(result['duplicate_details'])} duplicate keys found in Job List")
                    with st.expander("View Duplicate Keys"):
                        st.dataframe(result['duplicate_details'], use_container_width=True)
                else:
                    st.success("✅ No duplicate keys found in Job List")
                
                # Download section
                st.subheader("💾 Downloads")
                
                # Extract vessel name from job status filename
                vessel_name = "Unknown Vessel"
                if hasattr(st.session_state, 'job_status_file') and st.session_state.job_status_file:
                    vessel_name = extract_vessel_name(st.session_state.job_status_file.name)
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    # Excel export - primary download
                    excel_data = create_excel_export(result['enriched_data'], vessel_name)
                    st.download_button(
                        label="📊 Download Excel Report",
                        data=excel_data,
                        file_name=f"{vessel_name}_Job_List_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        help="Formatted Excel file with vessel name header and proper column ordering"
                    )
                
                with col2:
                    # Enriched data download (CSV)
                    st.download_button(
                        label="📥 Download CSV Data",
                        data=convert_df_to_csv(result['enriched_data']),
                        file_name="enriched_job_status.csv",
                        mime="text/csv"
                    )
                
                with col3:
                    # Unmatched rows download
                    if len(result['unmatched_details']) > 0:
                        st.download_button(
                            label="📥 Download Unmatched",
                            data=convert_df_to_csv(result['unmatched_details']),
                            file_name="unmatched_rows.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No unmatched rows")
                
                with col4:
                    # Fuzzy suggestions download
                    if len(result['fuzzy_suggestions']) > 0:
                        st.download_button(
                            label="💡 Download Suggestions",
                            data=convert_df_to_csv(result['fuzzy_suggestions']),
                            file_name="fuzzy_match_suggestions.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No suggestions")
                
            else:
                st.error(f"❌ Merge failed: {result['error']}")

if __name__ == "__main__":
    main()
