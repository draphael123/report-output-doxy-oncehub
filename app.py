"""
Weekly Report Generator - Local Flask Application
Upload spreadsheet files and generate consolidated Excel reports.
"""

from flask import Flask, render_template, request, send_file, flash, redirect, url_for, Response, jsonify, send_from_directory
import pandas as pd
from bs4 import BeautifulSoup
import re
import io
import tempfile
from datetime import datetime
import logging
import yaml
import os
from datetime import timedelta

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__, static_folder='static', static_url_path='/static')
app.secret_key = 'weekly-report-generator-secret-key-2026'

# Configuration loading functions
def load_config(config_path='config.yaml'):
    """Load configuration from YAML file"""
    if not os.path.exists(config_path):
        # Create default config if it doesn't exist
        default_config = get_default_config()
        save_config(default_config, config_path)
        return default_config
    
    try:
        with open(config_path, 'r') as f:
            config = yaml.safe_load(f)
        logger.info(f"Loaded configuration from {config_path}")
        return config
    except Exception as e:
        logger.error(f"Error loading config: {e}")
        return get_default_config()

def get_default_config():
    """Return default configuration if config.yaml doesn't exist"""
    return {
        'gusto_mappings': {
    'alg care and consulting inc': 'Ashley Grout',
    'alg care and consulting inc.': 'Ashley Grout',
    'cch ventures, pllc': 'Catherine Herrington',
    'cch ventures pllc': 'Catherine Herrington',
    'elizabeth gloor': 'Liz Gloor',
    'megan ryan riffle': 'Megan Ryan-Riffle',
    'jacquelyn sexton': 'Jacquelyn Sexton, NP',
        },
        'visit_mappings': {
            'terray humphrey': 'Darius Humphrey',
            'tim mack': 'Timothy Mack',
        },
        'na_providers': [
    'Bill Carbonneau NP',
    'Tzvi Doron',
    'Doron Stember',
    'Lindsay Burden NP',
    'Terray Humphrey',
    'Summer Denny',
        ],
        'excluded_names': ['daniel raphael', 'dan raphael', 'draphael'],
        'quality_thresholds': {
            'max_percentage_worked': 110,
            'min_visits_ratio': 0.3,
            'trt_ratio_min': 0.40,
            'trt_ratio_max': 0.85,
            'min_avg_duration': 5,
            'max_avg_duration': 60
        },
        'visit_durations': {
            'trt': 20,
            'hrt': 20,
            'other': 20
        }
    }

def save_config(config, config_path='config.yaml'):
    """Save configuration to YAML file"""
    try:
        with open(config_path, 'w') as f:
            yaml.dump(config, f, default_flow_style=False, sort_keys=False)
        logger.info(f"Saved configuration to {config_path}")
        return True
    except Exception as e:
        logger.error(f"Error saving config: {e}")
        return False

# Load configuration
config = load_config()
GUSTO_NAME_MAPPINGS = config.get('gusto_mappings', {})
VISIT_NAME_MAPPINGS = config.get('visit_mappings', {})
PROVIDERS_NA_HOURS = config.get('na_providers', [])
EXCLUDED_NAMES = config.get('excluded_names', [])
QUALITY_THRESHOLDS = config.get('quality_thresholds', {})
VISIT_DURATIONS = config.get('visit_durations', {'trt': 20, 'hrt': 20, 'other': 20})

# Import fuzzy matching (optional - will fail gracefully if not installed)
try:
    from fuzzywuzzy import fuzz, process
    FUZZY_AVAILABLE = True
except ImportError:
    FUZZY_AVAILABLE = False
    logger.warning("fuzzywuzzy not available - fuzzy matching disabled")

# DiagnosticLog class for tracking issues during report generation
class DiagnosticLog:
    """Tracks issues and warnings during report generation"""
    def __init__(self):
        self.warnings = []
        self.errors = []
        self.info = []
        self.data_stats = {}
        
    def add_warning(self, message, context=None):
        self.warnings.append({'message': message, 'context': context})
        logger.warning(f"{message} | Context: {context}")
    
    def add_error(self, message, context=None):
        self.errors.append({'message': message, 'context': context})
        logger.error(f"{message} | Context: {context}")
    
    def add_info(self, message, context=None):
        self.info.append({'message': message, 'context': context})
        logger.info(f"{message} | Context: {context}")
    
    def add_stat(self, key, value):
        self.data_stats[key] = value
    
    def get_summary(self):
        return {
            'warnings': self.warnings,
            'errors': self.errors,
            'info': self.info,
            'stats': self.data_stats
}

# File validation config
FILE_CONFIGS = {
    'doxy_file': {
        'name': 'Doxy Report',
        'extensions': ['.csv', '.xls', '.xlsx'],
        'required_columns': ['Provider name', 'Duration'],
        'max_size_mb': 10
    },
    'account_file': {
        'name': 'Account Detail Report',
        'extensions': ['.csv', '.xls', '.xlsx'],
        'max_size_mb': 10
    },
    'gusto_file': {
        'name': 'Gusto Hours',
        'extensions': ['.csv', '.xls', '.xlsx'],
        'max_size_mb': 10
    },
    'booking_file': {
        'name': 'OnceHub Booking Summary',
        'extensions': ['.csv', '.xls', '.xlsx'],
        'required_columns': ['Booking page'],
        'max_size_mb': 10
    }
}


def should_exclude_name(name):
    """Check if a name should be excluded from reports."""
    if pd.isna(name):
        return False
    name_lower = str(name).lower().strip()
    for excluded in EXCLUDED_NAMES:
        if excluded in name_lower:
            return True
    return False


def filter_out_upload_day(doxy_df, diagnostic=None):
    """Process Doxy report data (includes all days including Sundays).
    """
    if doxy_df is None or doxy_df.empty:
        return doxy_df
    
    original_count = len(doxy_df)
    
    if 'Date' not in doxy_df.columns:
        if diagnostic:
            diagnostic.add_info("No Date column in Doxy report")
        return doxy_df
    
    # Parse dates to validate them (but don't filter any days)
    doxy_df = doxy_df.copy()
    doxy_df['Parsed_Date'] = pd.to_datetime(doxy_df['Date'], errors='coerce')
    
    # Count unparseable dates
    unparseable_count = doxy_df['Parsed_Date'].isna().sum()
    if unparseable_count > 0 and diagnostic:
        diagnostic.add_warning(f"{unparseable_count} dates in Doxy report could not be parsed", {
            'total_rows': original_count,
            'unparseable': unparseable_count
        })
    
    # Keep only rows with valid dates
    doxy_df = doxy_df[doxy_df['Parsed_Date'].notna()]
    
    # Drop the temporary column
    doxy_df = doxy_df.drop(columns=['Parsed_Date'], errors='ignore')
    
    final_count = len(doxy_df)
    
    if diagnostic:
        diagnostic.add_info("Processed Doxy report (all days included)", {
            'original_rows': original_count,
            'final_rows': final_count,
            'rows_with_invalid_dates': original_count - final_count
        })
    
    return doxy_df


def validate_file(file_obj, config):
    """Validate a file against its configuration."""
    errors = []
    
    # Check extension
    filename = file_obj.filename.lower()
    if not any(filename.endswith(ext) for ext in config['extensions']):
        errors.append(f"Invalid file type. Expected: {', '.join(config['extensions'])}")
    
    # Check file size
    file_obj.seek(0, 2)  # Seek to end
    size_mb = file_obj.tell() / (1024 * 1024)
    file_obj.seek(0)  # Reset to beginning
    
    if size_mb > config['max_size_mb']:
        errors.append(f"File too large ({size_mb:.1f}MB). Max: {config['max_size_mb']}MB")
    
    if size_mb == 0:
        errors.append("File is empty")
    
    return errors


def parse_duration_to_minutes(duration_str):
    """Convert duration string (HH:MM:SS) to minutes."""
    if pd.isna(duration_str) or duration_str == "No data" or not duration_str:
        return None
    try:
        parts = str(duration_str).split(":")
        if len(parts) == 3:
            hours, minutes, seconds = map(int, parts)
            return hours * 60 + minutes + seconds / 60
        return None
    except (ValueError, AttributeError):
        return None


def get_doxy_visits(doxy_df, diagnostic=None):
    """Section 1: Count visits per provider from Doxy Report."""
    if diagnostic:
        original_count = len(doxy_df) if doxy_df is not None and not doxy_df.empty else 0
        diagnostic.add_info("Processing Doxy visits", {'input_rows': original_count})
    
    # Process Doxy data (validate dates, all days included)
    doxy_df = filter_out_upload_day(doxy_df, diagnostic=diagnostic)
    
    # Filter out excluded names
    if doxy_df is not None and not doxy_df.empty:
        excluded_mask = doxy_df['Provider name'].apply(should_exclude_name)
        excluded_count = excluded_mask.sum()
        if excluded_count > 0:
            excluded_names = doxy_df[excluded_mask]['Provider name'].unique().tolist()
            if diagnostic:
                diagnostic.add_info(f"Excluded {excluded_count} visits from excluded providers", {
                    'excluded_providers': excluded_names,
                    'excluded_count': excluded_count
                })
        doxy_df = doxy_df[~excluded_mask]
    
    if doxy_df is None or doxy_df.empty:
        if diagnostic:
            diagnostic.add_warning("Doxy DataFrame is empty after filtering")
        return pd.DataFrame(columns=['Provider name', 'Total Visits'])
    
    visits = doxy_df.groupby("Provider name").size().reset_index(name="Total Visits")
    visits = visits.sort_values("Provider name", ascending=True)
    
    if diagnostic:
        diagnostic.add_stat('doxy_providers_count', len(visits))
        diagnostic.add_stat('doxy_total_visits', int(visits['Total Visits'].sum()))
        diagnostic.add_info("Completed Doxy visits processing", {
            'providers': len(visits),
            'total_visits': int(visits['Total Visits'].sum())
        })
    
    return visits


def get_oncehub_visits(booking_df, diagnostic=None):
    """Section 2: Get visit counts from OnceHub Booking Summary."""
    if booking_df is None or booking_df.empty:
        if diagnostic:
            diagnostic.add_info("OnceHub booking data not provided or empty")
        return None
    
    if diagnostic:
        diagnostic.add_info("Processing OnceHub visits", {'input_rows': len(booking_df)})
    
    if 'Booking page' not in booking_df.columns:
        if diagnostic:
            diagnostic.add_warning("OnceHub file missing 'Booking page' column", {
                'available_columns': list(booking_df.columns)
            })
        return None
    
    booking_df = booking_df.copy()
    booking_df['Provider'] = booking_df['Booking page'].str.replace(r'\s*\([^)]*\)', '', regex=True).str.strip()
    
    # Filter out excluded names
    excluded_mask = booking_df['Provider'].apply(should_exclude_name)
    excluded_count = excluded_mask.sum()
    if excluded_count > 0:
        excluded_names = booking_df[excluded_mask]['Provider'].unique().tolist()
        if diagnostic:
            diagnostic.add_info(f"Excluded {excluded_count} OnceHub entries from excluded providers", {
                'excluded_providers': excluded_names
            })
    booking_df = booking_df[~excluded_mask]
    
    required_cols = ['All activities', 'Scheduled', 'Completed', 'Canceled', 'No-show']
    missing_cols = [col for col in required_cols if col not in booking_df.columns]
    if missing_cols and diagnostic:
        diagnostic.add_warning(f"OnceHub file missing columns: {missing_cols}", {
            'available_columns': list(booking_df.columns)
        })
    
    result = booking_df[['Provider'] + [col for col in required_cols if col in booking_df.columns]].copy()
    # Use Completed as the primary count (Total Visits)
    if 'Completed' in result.columns:
        result['Total Visits'] = result['Completed']
    else:
        result['Total Visits'] = 0
        if diagnostic:
            diagnostic.add_warning("OnceHub file missing 'Completed' column, setting Total Visits to 0")
    
    result = result[['Provider', 'Total Visits'] + [col for col in required_cols if col in result.columns]].copy()
    result = result.sort_values('Provider', ascending=True)
    
    if diagnostic:
        diagnostic.add_stat('oncehub_providers_count', len(result))
        diagnostic.add_stat('oncehub_total_visits', int(result['Total Visits'].sum()) if 'Total Visits' in result.columns else 0)
        diagnostic.add_info("Completed OnceHub visits processing", {
            'providers': len(result),
            'total_visits': int(result['Total Visits'].sum()) if 'Total Visits' in result.columns else 0
        })
    
    return result


def get_visits_by_program(account_content, is_csv=False, start_date=None, end_date=None, diagnostic=None):
    """Section 3: Parse AccountDetailReport and categorize visits.
    
    Args:
        account_content: Content of Account Detail Report (string)
        is_csv: Whether content is CSV format
        start_date: Optional start date (datetime) to filter visits
        end_date: Optional end date (datetime) to filter visits
        diagnostic: Optional DiagnosticLog instance
    """
    if diagnostic:
        diagnostic.add_info("Starting visits by program processing", {
            'is_csv': is_csv,
            'has_start_date': start_date is not None,
            'has_end_date': end_date is not None
        })
    
    if is_csv:
        # Parse as CSV
        try:
            df = pd.read_csv(io.StringIO(account_content))
            if diagnostic:
                diagnostic.add_info("Successfully parsed CSV", {
                    'rows': len(df),
                    'columns': list(df.columns),
                    'sample_data': df.head(3).to_dict('records') if len(df) > 0 else []
                })
            # Check if DataFrame is actually empty
            if df.empty:
                if diagnostic:
                    diagnostic.add_warning("CSV parsed but DataFrame is empty - check file content")
                return pd.DataFrame(columns=['Provider', 'TRT', 'HRT', 'Other', 'Total'])
        except Exception as e:
            # If CSV parsing fails, return empty DataFrame
            if diagnostic:
                diagnostic.add_error("CSV parsing failed", {'error': str(e), 'content_length': len(account_content) if account_content else 0})
            return pd.DataFrame(columns=['Provider', 'TRT', 'HRT', 'Other', 'Total'])
        
        # Find date column BEFORE renaming (for date filtering)
        original_date_col = None
        for col in df.columns:
            col_lower = str(col).lower()
            if 'date' in col_lower and 'meeting' in col_lower:
                original_date_col = col
                break
            elif 'date' in col_lower and original_date_col is None:
                original_date_col = col
        
        # Store the original date column name before any renaming
        date_col_for_filtering = original_date_col
        
        # Map columns - adjust based on your CSV structure
        # Expected columns: Status, Owner/Provider, Event Type
        # IMPORTANT: Exclude date columns from being mapped to Provider
        col_mapping = {}
        for col in df.columns:
            col_lower = str(col).lower().strip()
            # Skip date columns entirely
            if 'date' in col_lower:
                continue
            # Map Status column
            if 'status' in col_lower:
                col_mapping['Status'] = col
            # Map Provider/Owner column
            elif 'owner' in col_lower or ('provider' in col_lower and 'page' not in col_lower):
                col_mapping['Provider'] = col
            # Map Event Type column - prioritize "event type name" over "subject/event type"
            elif 'event' in col_lower and 'type' in col_lower and 'name' in col_lower:
                # Prefer "Event type name" over "Subject/Event type"
                if 'Event Type' not in col_mapping or 'name' in col_mapping.get('Event Type', '').lower():
                    col_mapping['Event Type'] = col
            elif 'event' in col_lower and 'type' in col_lower:
                # Map "Subject/Event type" or similar if we haven't found a better one
                if 'Event Type' not in col_mapping:
                    col_mapping['Event Type'] = col
            # Also check for "subject/event type" or similar
            elif 'subject' in col_lower and 'event' in col_lower:
                if 'Event Type' not in col_mapping:
                    col_mapping['Event Type'] = col
        
        if diagnostic:
            diagnostic.add_info("CSV column detection", {
                'date_column': original_date_col,
                'mapped_columns': col_mapping,
                'all_columns': list(df.columns)
            })
        
        # Rename columns
        if col_mapping:
            df = df.rename(columns={v: k for k, v in col_mapping.items()})
        
        # Preserve date column for filtering - keep it with original name if it wasn't renamed
        if original_date_col and original_date_col not in col_mapping.values():
            # Date column wasn't renamed, so it still has its original name
            # Make sure it's still in the dataframe
            if original_date_col not in df.columns:
                # If it was somehow lost, we need to find it again
                for col in df.columns:
                    if 'date' in str(col).lower() and 'meeting' in str(col).lower():
                        original_date_col = col
                        break
        
        # Ensure required columns exist
        if 'Status' not in df.columns:
            # Try to find status column with different names
            for col in df.columns:
                col_lower = str(col).lower()
                if 'status' in col_lower or 'state' in col_lower:
                    df['Status'] = df[col]
                    break
            if 'Status' not in df.columns:
                if diagnostic:
                    diagnostic.add_warning("No Status column found, assuming all Completed")
                df['Status'] = 'Completed'
        
        if 'Provider' not in df.columns:
            # Try to find a name-like column with more variations
            # IMPORTANT: Exclude date columns
            for col in df.columns:
                col_lower = str(col).lower()
                # Skip date columns
                if 'date' in col_lower:
                    continue
                if any(term in col_lower for term in ['name', 'owner', 'provider', 'assigned', 'contact']):
                    df['Provider'] = df[col]
                    break
            if 'Provider' not in df.columns and len(df.columns) > 0:
                # Use first non-date column as fallback
                for col in df.columns:
                    col_lower = str(col).lower()
                    if 'date' not in col_lower and col not in ['Status', 'Event Type']:
                        df['Provider'] = df[col]
                        break
        
        if 'Event Type' not in df.columns:
            # Try to find event type column with different names
            for col in df.columns:
                col_lower = str(col).lower()
                if 'event' in col_lower or 'program' in col_lower or 'category' in col_lower:
                    df['Event Type'] = df[col]
                    break
            if 'Event Type' not in df.columns:
                if diagnostic:
                    diagnostic.add_warning("No Event Type column found, defaulting to 'Other'")
                df['Event Type'] = 'Other'
    else:
        # Parse as HTML (XLS files from OnceHub are actually HTML)
        if diagnostic:
            diagnostic.add_info("Parsing HTML content")
        soup = BeautifulSoup(account_content, 'html.parser')
        rows = soup.find_all('tr')
        
        data = []
        skipped_header_rows = 0
        for row in rows:
            cells = row.find_all('td')
            if len(cells) >= 7:
                first_cell = cells[0]
                if first_cell.get('style') and 'border-style:solid' in first_cell.get('style', ''):
                    status = cells[3].get_text(strip=True)
                    owner = cells[5].get_text(strip=True)
                    event_type = cells[6].get_text(strip=True)
                    
                    # Skip header row (where Status is "Status" or Owner is "Booking page owner")
                    if status.lower() == 'status' or owner.lower() == 'booking page owner':
                        skipped_header_rows += 1
                        continue
                    
                    data.append({
                        'Status': status,
                        'Provider': owner,
                        'Event Type': event_type
                    })
        
        if diagnostic:
            diagnostic.add_info("HTML parsing completed", {
                'total_rows_found': len(rows),
                'data_rows': len(data),
                'skipped_header_rows': skipped_header_rows
                    })
        
        df = pd.DataFrame(data)
    
    # Check if DataFrame is empty BEFORE column mapping/renaming
    if df.empty:
        if diagnostic:
            diagnostic.add_warning("Account Detail Report DataFrame is empty after parsing")
        return pd.DataFrame(columns=['Provider', 'TRT', 'HRT', 'Other', 'Total'])
    
    if diagnostic:
        diagnostic.add_info("DataFrame after initial parsing", {
            'rows': len(df),
            'columns': list(df.columns)
        })
    
    # Ensure Provider column exists
    if 'Provider' not in df.columns:
        if diagnostic:
            diagnostic.add_error("Provider column not found after column mapping", {
                'available_columns': list(df.columns)
            })
        return pd.DataFrame(columns=['Provider', 'TRT', 'HRT', 'Other', 'Total'])
    
    # Find date column for filtering
    # Use the stored date column name, or search in current columns
    date_col = None
    if 'date_col_for_filtering' in locals() and date_col_for_filtering:
        # Check if the original date column still exists (it wasn't renamed)
        if date_col_for_filtering in df.columns:
            date_col = date_col_for_filtering
        else:
            # Date column was renamed or lost, search for it
            date_col = None
    
    if not date_col:
        # Fallback: search in current columns (after renaming)
        for col in df.columns:
            col_lower = str(col).lower()
            if 'date' in col_lower and 'meeting' in col_lower:
                date_col = col
                break
            elif 'date' in col_lower:
                date_col = col
                break
    
    # Parse dates if we have a date column
    if date_col and date_col in df.columns:
        # Parse dates from the date column
        def parse_date(date_str):
            if pd.isna(date_str):
                return None
            try:
                date_str = str(date_str)
                # Try parsing formats like "Sun, Feb 1, 2026, 08:00 AM - 08:20 AM" or "Fri, Dec 26, 2025, 07:00 AM - 07:20 AM"
                parts = date_str.split(',')
                if len(parts) >= 3:
                    # Take "Feb 1, 2026" or "Dec 26, 2025"
                    date_part = ','.join(parts[1:3]).strip()
                    # Try with day without leading zero first
                    parsed = pd.to_datetime(date_part, format='%b %d, %Y', errors='coerce')
                    if pd.isna(parsed):
                        # Try alternative format
                        parsed = pd.to_datetime(date_part, errors='coerce')
                    return parsed
                # Fallback to full string parsing
                return pd.to_datetime(date_str, errors='coerce')
            except:
                return pd.to_datetime(date_str, errors='coerce')
        
        df['Parsed_Date'] = df[date_col].apply(parse_date)
        
        # Count unparseable dates
        unparseable = df['Parsed_Date'].isna().sum()
        if unparseable > 0 and diagnostic:
            if date_col and date_col in df.columns:
                sample_unparseable = df[df['Parsed_Date'].isna()][date_col].head(5).tolist()
            else:
                sample_unparseable = []
            diagnostic.add_warning(f"{unparseable} dates in Account Detail could not be parsed", {
                'sample_dates': sample_unparseable,
                'total_rows': len(df)
            })
        
        # Keep only rows with valid parsed dates (all days including Sundays are included)
        valid_dates = df['Parsed_Date'].notna()
        df_with_dates = df[valid_dates].copy()
        if len(df_with_dates) > 0:
            df = df_with_dates
        else:
            # If no dates could be parsed, keep all rows (they'll be filtered by date range if provided)
            df = df.copy()
            if diagnostic:
                diagnostic.add_warning("No dates could be parsed from Account Detail Report - skipping date filtering")
        
        if diagnostic:
            diagnostic.add_info("Processed Account Detail dates (all days included)", {
                'rows_with_valid_dates': len(df)
            })
        
        # Filter by date range if dates are provided
        if (start_date or end_date) and 'Parsed_Date' in df.columns:
            before_date_filter = len(df)
            # Only filter if we have valid parsed dates
            valid_dates_mask = df['Parsed_Date'].notna()
            df_valid = df[valid_dates_mask].copy()
            
            if len(df_valid) > 0:
                # Convert start_date and end_date to datetime if they're strings
                if start_date:
                    if isinstance(start_date, str):
                        start_date = pd.to_datetime(start_date)
                    df_valid = df_valid[df_valid['Parsed_Date'] >= start_date]
                
                if end_date:
                    if isinstance(end_date, str):
                        end_date = pd.to_datetime(end_date)
                    # Add one day to include the end date
                    end_date_inclusive = end_date + pd.Timedelta(days=1)
                    df_valid = df_valid[df_valid['Parsed_Date'] < end_date_inclusive]
                
                df = df_valid
            else:
                # If no valid dates, keep rows without dates (they won't be filtered)
                if diagnostic:
                    diagnostic.add_warning("No valid dates found for date range filtering - keeping all rows")
            
            if diagnostic:
                diagnostic.add_info("Applied date range filter", {
                    'start_date': str(start_date) if start_date else None,
                    'end_date': str(end_date) if end_date else None,
                    'before_filter': before_date_filter,
                    'after_filter': len(df),
                    'rows_removed': before_date_filter - len(df)
                })
        
        # Drop the temporary column and the original date column (if it still exists)
        df = df.drop(columns=['Parsed_Date'], errors='ignore')
        if date_col and date_col in df.columns and date_col != 'Provider':
            df = df.drop(columns=[date_col], errors='ignore')
    
    # Drop the original date column if it still exists (to prevent it from being used as Provider)
    if date_col and date_col in df.columns and date_col != 'Provider':
        df = df.drop(columns=[date_col], errors='ignore')
    
    # Verify Provider column exists and is not a date column
    if 'Provider' not in df.columns:
        # Try to find a name-like column, but exclude date columns
        for col in df.columns:
            col_lower = str(col).lower()
            # Skip date columns
            if 'date' in col_lower:
                continue
            if any(term in col_lower for term in ['name', 'owner', 'provider', 'assigned', 'contact']):
                df['Provider'] = df[col]
                break
        # If still not found, use first non-date column
        if 'Provider' not in df.columns:
            for col in df.columns:
                col_lower = str(col).lower()
                if 'date' not in col_lower and col not in ['Status', 'Event Type', 'Category']:
                    df['Provider'] = df[col]
                    break
    
    # Final check - if Provider column contains dates, we have a problem
    if 'Provider' in df.columns and len(df) > 0:
        # Check if Provider values look like dates (format: MM/DD/YYYY or similar)
        sample_provider = str(df['Provider'].iloc[0]) if len(df) > 0 else ''
        # Check if it looks like a date (contains / and starts with digits)
        if '/' in sample_provider and len(sample_provider) > 8 and any(char.isdigit() for char in sample_provider[:2]):
            # This looks like a date, try to find the real Provider column
            for col in df.columns:
                col_lower = str(col).lower()
                if 'date' not in col_lower and col != 'Provider' and col not in ['Status', 'Event Type', 'Category', 'Parsed_Date']:
                    # Check if this column has name-like values (not dates)
                    if len(df) > 0:
                        sample_val = str(df[col].iloc[0])
                        # If it doesn't look like a date, use it as Provider
                        if not ('/' in sample_val and len(sample_val) > 8 and any(char.isdigit() for char in sample_val[:2])):
                            df['Provider'] = df[col]
                            break
    
    # Filter by Status only if Status column exists
    if 'Status' in df.columns:
        before_status_filter = len(df)
        df = df[df['Status'] == 'Completed']
        if diagnostic:
            diagnostic.add_info("Filtered for Completed status", {
                'before_filter': before_status_filter,
                'after_filter': len(df),
                'rows_removed': before_status_filter - len(df)
            })
    
    # Filter out excluded names
    before_exclusion = len(df)
    excluded_mask = df['Provider'].apply(should_exclude_name)
    excluded_count = excluded_mask.sum()
    if excluded_count > 0:
        excluded_names = df[excluded_mask]['Provider'].unique().tolist()
        if diagnostic:
            diagnostic.add_info(f"Excluded {excluded_count} visits from excluded providers", {
                'excluded_providers': excluded_names
            })
    df = df[~excluded_mask]
    
    if df.empty:
        if diagnostic:
            diagnostic.add_warning("Account Detail DataFrame is empty after all filtering")
        return pd.DataFrame(columns=['Provider', 'TRT', 'HRT', 'Other', 'Total'])
    
    def get_category(event_type):
        if pd.isna(event_type):
            return 'Other'
        event_upper = str(event_type).upper()
        if 'TRT' in event_upper or 'FOUNTAINTRT' in event_upper:
            return 'TRT'
        elif 'HRT' in event_upper:
            return 'HRT'
        else:
            return 'Other'
    
    df['Category'] = df['Event Type'].apply(get_category)
    
    if df.empty:
        return pd.DataFrame(columns=['Provider', 'TRT', 'HRT', 'Other', 'Total'])
    
    # Ensure we only have Provider, Status, Event Type, and Category columns before pivot
    df = df[['Provider', 'Category']].copy()
    
    pivot = df.pivot_table(
        index='Provider',
        columns='Category',
        aggfunc='size',
        fill_value=0
    )
    
    if pivot.empty:
        return pd.DataFrame(columns=['Provider', 'TRT', 'HRT', 'Other', 'Total'])
    
    pivot = pivot.reset_index()
    
    cols_order = ['Provider']
    for col in ['TRT', 'HRT', 'Other']:
        if col in pivot.columns:
            cols_order.append(col)
        else:
            pivot[col] = 0
    
    # Ensure all required columns exist
    for col in ['TRT', 'HRT', 'Other']:
        if col not in pivot.columns:
            pivot[col] = 0
    
    pivot = pivot[cols_order]
    
    numeric_cols = [col for col in pivot.columns if col != 'Provider']
    pivot['Total'] = pivot[numeric_cols].sum(axis=1)
    pivot = pivot.sort_values('Provider', ascending=True)
    
    if diagnostic:
        diagnostic.add_stat('visits_by_program_providers', len(pivot))
        diagnostic.add_stat('visits_by_program_trt', int(pivot['TRT'].sum()) if 'TRT' in pivot.columns else 0)
        diagnostic.add_stat('visits_by_program_hrt', int(pivot['HRT'].sum()) if 'HRT' in pivot.columns else 0)
        diagnostic.add_stat('visits_by_program_other', int(pivot['Other'].sum()) if 'Other' in pivot.columns else 0)
        diagnostic.add_info("Completed visits by program processing", {
            'providers': len(pivot),
            'trt_total': int(pivot['TRT'].sum()) if 'TRT' in pivot.columns else 0,
            'hrt_total': int(pivot['HRT'].sum()) if 'HRT' in pivot.columns else 0,
            'other_total': int(pivot['Other'].sum()) if 'Other' in pivot.columns else 0,
            'total_visits': int(pivot['Total'].sum()) if 'Total' in pivot.columns else 0
        })
    
    return pivot


def get_gusto_hours(gusto_df, doxy_providers, diagnostic=None):
    """Section 4: Extract Gusto hours for providers in visit data."""
    if diagnostic:
        diagnostic.add_info("Processing Gusto hours", {
            'input_rows': len(gusto_df) if gusto_df is not None and not gusto_df.empty else 0,
            'doxy_providers_count': len(doxy_providers) if doxy_providers else 0
        })
    
    if gusto_df is None or gusto_df.empty:
        if diagnostic:
            diagnostic.add_warning("Gusto DataFrame is empty")
        return pd.DataFrame(columns=['Name', 'Total hours'])
    
    if len(gusto_df.columns) >= 4:
        gusto_df.columns = ['Name', 'Title', 'Manager', 'Total hours'] + list(gusto_df.columns[4:])
    
    gusto_df['Name'] = gusto_df['Name'].astype(str).str.strip().str.replace('"', '')
    
    # Apply name mappings (company names -> provider names)
    def apply_name_mapping(name):
        if pd.isna(name):
            return name
        name_lower = str(name).lower().strip()
        return GUSTO_NAME_MAPPINGS.get(name_lower, name)
    
    original_names = gusto_df['Name'].copy()
    gusto_df['Name'] = gusto_df['Name'].apply(apply_name_mapping)
    
    # Track which names were mapped
    mapped_count = (original_names != gusto_df['Name']).sum()
    if mapped_count > 0 and diagnostic:
        mapped_names = gusto_df[original_names != gusto_df['Name']][['Name']].drop_duplicates()['Name'].tolist()
        diagnostic.add_info(f"Applied name mappings to {mapped_count} Gusto entries", {
            'mapped_providers': mapped_names[:10]  # Show first 10
        })
    
    def normalize_name(name):
        if pd.isna(name):
            return ''
        name = str(name).strip()
        name = re.sub(r'\s+(NP|FNP-C|MD|PA|LLC|Inc\.?|INC\.?|PLLC)$', '', name, flags=re.IGNORECASE)
        name = re.sub(r',\s*NP$', '', name, flags=re.IGNORECASE)
        return name.lower().strip()
    
    doxy_normalized = set(normalize_name(p) for p in doxy_providers)
    gusto_df['Name_normalized'] = gusto_df['Name'].apply(normalize_name)
    
    def is_in_doxy(name_normalized):
        if not name_normalized:
            return False
        for doxy_name in doxy_normalized:
            name_parts = set(name_normalized.split())
            doxy_parts = set(doxy_name.split())
            if len(name_parts.intersection(doxy_parts)) >= 2:
                return True
            if name_normalized == doxy_name:
                return True
        return False
    
    gusto_df['In_Doxy'] = gusto_df['Name_normalized'].apply(is_in_doxy)
    
    # Track matching
    matched_providers = gusto_df[gusto_df['In_Doxy']]['Name'].unique().tolist()
    unmatched_gusto = gusto_df[~gusto_df['In_Doxy']]['Name'].unique().tolist()
    
    if diagnostic:
        diagnostic.add_info("Gusto provider matching", {
            'matched': len(matched_providers),
            'unmatched': len(unmatched_gusto),
            'unmatched_providers': unmatched_gusto[:10]  # Show first 10
        })
        if len(unmatched_gusto) > 0:
            diagnostic.add_warning(f"{len(unmatched_gusto)} Gusto providers had no visit matches", {
                'unmatched_providers': unmatched_gusto[:10]
            })
    
    filtered = gusto_df[gusto_df['In_Doxy']][['Name', 'Total hours']].copy()
    filtered['Total hours'] = pd.to_numeric(filtered['Total hours'], errors='coerce')
    
    # Track invalid hours
    invalid_hours_count = filtered['Total hours'].isna().sum()
    if invalid_hours_count > 0 and diagnostic:
        diagnostic.add_warning(f"{invalid_hours_count} Gusto entries have invalid hours", {
            'total_entries': len(filtered)
        })
    
    filtered = filtered[filtered['Total hours'] > 0]
    
    # Filter out excluded names
    excluded_mask = filtered['Name'].apply(should_exclude_name)
    excluded_count = excluded_mask.sum()
    if excluded_count > 0 and diagnostic:
        excluded_names = filtered[excluded_mask]['Name'].unique().tolist()
        diagnostic.add_info(f"Excluded {excluded_count} Gusto entries from excluded providers", {
            'excluded_providers': excluded_names
        })
    filtered = filtered[~excluded_mask]
    
    filtered = filtered.sort_values('Name', ascending=True)
    
    # Add providers with N/A hours
    na_providers = pd.DataFrame({
        'Name': PROVIDERS_NA_HOURS,
        'Total hours': ['N/A'] * len(PROVIDERS_NA_HOURS)
    })
    filtered = pd.concat([filtered, na_providers], ignore_index=True)
    filtered = filtered.sort_values('Name', ascending=True)
    
    if diagnostic:
        providers_with_hours = len(filtered[filtered['Total hours'] != 'N/A'])
        diagnostic.add_stat('gusto_providers_with_hours', providers_with_hours)
        diagnostic.add_info("Completed Gusto hours processing", {
            'total_providers': len(filtered),
            'providers_with_hours': providers_with_hours,
            'providers_na_hours': len(filtered[filtered['Total hours'] == 'N/A'])
        })
    
    return filtered


def get_doxy_performance_metrics(doxy_df, diagnostic=None):
    """Section 5: Calculate performance metrics from Doxy Report."""
    if diagnostic:
        diagnostic.add_info("Processing Doxy performance metrics")
    
    # Filter out the upload day (Sunday) to show only 6 days
    doxy_df = filter_out_upload_day(doxy_df, diagnostic=diagnostic)
    # Filter out excluded names
    doxy_df = doxy_df[~doxy_df['Provider name'].apply(should_exclude_name)]
    
    doxy_df['Duration_Minutes'] = doxy_df['Duration'].apply(parse_duration_to_minutes)
    df_valid = doxy_df[doxy_df['Duration_Minutes'].notna()].copy()
    
    metrics = df_valid.groupby('Provider name').agg(
        Total_Visits=('Duration_Minutes', 'count'),
        Visits_Over_20_Min=('Duration_Minutes', lambda x: (x > 20).sum()),
        Hours_Over_20_Min=('Duration_Minutes', lambda x: (x[x > 20].sum() / 60)),
        Avg_Duration_Min=('Duration_Minutes', 'mean')
    ).reset_index()
    
    metrics['Pct_Over_20_Min'] = (metrics['Visits_Over_20_Min'] / metrics['Total_Visits'] * 100).round(1)
    metrics['Avg_Duration_Min'] = metrics['Avg_Duration_Min'].round(2)
    metrics['Hours_Over_20_Min'] = metrics['Hours_Over_20_Min'].round(2)
    
    metrics.columns = ['Provider', 'Total Visits', 'Visits Over 20 Min', 
                       'Hours on 20+ Min Visits', 'Avg Duration (min)', '% Over 20 Min']
    
    metrics = metrics[['Provider', 'Total Visits', 'Visits Over 20 Min', 
                       '% Over 20 Min', 'Hours on 20+ Min Visits', 'Avg Duration (min)']]
    
    metrics = metrics.sort_values('Provider', ascending=True)
    
    return metrics


def get_hours_worked(gusto_hours, visits_by_program, diagnostic=None):
    """Section 6: Calculate hours worked based on visit types.
    TRT visits = 20 minutes, HRT visits = 20 minutes.
    """
    if diagnostic:
        diagnostic.add_info("Processing hours worked calculation", {
            'gusto_providers': len(gusto_hours) if gusto_hours is not None and not gusto_hours.empty else 0,
            'visits_providers': len(visits_by_program) if visits_by_program is not None and not visits_by_program.empty else 0
        })
    
    def normalize_name(name):
        if pd.isna(name):
            return ''
        name = str(name).strip().lower()
        name = re.sub(r'\s+(np|fnp-c|md|pa|llc|inc\.?|pllc)$', '', name, flags=re.IGNORECASE)
        name = re.sub(r',?\s*np$', '', name, flags=re.IGNORECASE)  # Handle ", NP" or " NP"
        name = name.rstrip(',').strip()  # Remove any trailing commas
        return name.strip()
    
    # Filter out N/A hours for the calculation
    gusto = gusto_hours[gusto_hours['Total hours'] != 'N/A'].copy()
    gusto['Total hours'] = pd.to_numeric(gusto['Total hours'], errors='coerce')
    visits = visits_by_program.copy()
    
    # Initialize result DataFrame
    result = pd.DataFrame(columns=['Provider', 'Gusto Hours', 'TRT Visits (20 min)', 'HRT Visits (20 min)', 'Total Visits', 'Hours Worked', 'Percentage Worked'])
    
    # Handle empty visits DataFrame - still process providers with Gusto hours
    if visits.empty or 'Provider' not in visits.columns:
        # Add providers with Gusto hours (even if no visits)
        if not gusto.empty and 'Name' in gusto.columns:
            gusto['Name_norm'] = gusto['Name'].apply(normalize_name)
            for _, row in gusto.iterrows():
                new_row = pd.DataFrame({
                    'Provider': [row['Name']],
                    'Gusto Hours': [row['Total hours']],
                    'TRT Visits (20 min)': [0],
                    'HRT Visits (20 min)': [0],
                    'Total Visits': [0],
                    'Hours Worked': [0],
                    'Percentage Worked': [0]
                })
                result = pd.concat([result, new_row], ignore_index=True)
        
        # Add N/A providers with 0 hours
        for provider_name in PROVIDERS_NA_HOURS:
            new_row = pd.DataFrame({
                'Provider': [provider_name],
                'Gusto Hours': ['N/A'],
                'TRT Visits (20 min)': [0],
                'HRT Visits (20 min)': [0],
                'Total Visits': [0],
                'Hours Worked': [0],
                'Percentage Worked': ['N/A']
            })
            result = pd.concat([result, new_row], ignore_index=True)
        
        if not result.empty:
            result = result.sort_values('Provider', ascending=True)
        return result
    
    # Handle empty gusto DataFrame
    if gusto.empty or 'Name' not in gusto.columns:
        result = pd.DataFrame(columns=['Provider', 'Gusto Hours', 'TRT Visits (20 min)', 'HRT Visits (20 min)', 'Total Visits', 'Hours Worked'])
        # Add N/A providers with calculated hours from visits
        visits['Name_norm'] = visits['Provider'].apply(normalize_name)
        for provider_name in PROVIDERS_NA_HOURS:
            provider_norm = normalize_name(provider_name)
            lookup_name = VISIT_NAME_MAPPINGS.get(provider_norm, provider_norm)
            if lookup_name != provider_norm:
                lookup_name = normalize_name(lookup_name)
            visit_match = visits[visits['Name_norm'] == lookup_name]
            if not visit_match.empty:
                row = visit_match.iloc[0]
                trt = row['TRT'] if 'TRT' in row else 0
                hrt = row['HRT'] if 'HRT' in row else 0
                other = row['Other'] if 'Other' in row else 0
                total = row['Total'] if 'Total' in row else (trt + hrt + other)
            else:
                trt, hrt, other, total = 0, 0, 0, 0
            trt_duration = VISIT_DURATIONS.get('trt', 20)
            hrt_duration = VISIT_DURATIONS.get('hrt', 20)
            other_duration = VISIT_DURATIONS.get('other', 20)
            hours_worked = round((trt * trt_duration + hrt * hrt_duration + other * other_duration) / 60, 2)
            new_row = pd.DataFrame({
                'Provider': [provider_name],
                'Gusto Hours': ['N/A'],
                'TRT Visits (20 min)': [trt],
                'HRT Visits (20 min)': [hrt],
                'Total Visits': [total],
                'Hours Worked': [hours_worked],
                'Percentage Worked': ['N/A']
            })
            result = pd.concat([result, new_row], ignore_index=True)
        return result
    
    gusto['Name_norm'] = gusto['Name'].apply(normalize_name)
    
    # Use left join to include all providers with Gusto hours, even if they have no visits
    if not visits.empty and 'Provider' in visits.columns:
        # Apply name mappings to visits before normalizing
        def apply_visit_mapping(name):
            if pd.isna(name):
                return name
            name_lower = str(name).lower().strip()
            return VISIT_NAME_MAPPINGS.get(name_lower, name)
        
        visits['Provider_mapped'] = visits['Provider'].apply(apply_visit_mapping)
        visits['Name_norm'] = visits['Provider_mapped'].apply(normalize_name)
        merged = pd.merge(gusto, visits, on='Name_norm', how='left')
    else:
        # If visits is empty, create a merged DataFrame with just Gusto data
        merged = gusto.copy()
        merged['TRT'] = 0
        merged['HRT'] = 0
        merged['Other'] = 0
        merged['Total'] = 0
    
    # Initialize result DataFrame with proper columns
    if merged.empty:
        result = pd.DataFrame(columns=['Provider', 'Gusto Hours', 'TRT Visits (20 min)', 'HRT Visits (20 min)', 'Total Visits', 'Hours Worked', 'Percentage Worked'])
    else:
        # Get TRT and HRT counts (default to 0 if column doesn't exist)
        merged['TRT Visits'] = merged['TRT'].fillna(0) if 'TRT' in merged.columns else 0
        merged['HRT Visits'] = merged['HRT'].fillna(0) if 'HRT' in merged.columns else 0
        merged['Other Visits'] = merged['Other'].fillna(0) if 'Other' in merged.columns else 0
        merged['Total Visits'] = merged['Total'].fillna(0) if 'Total' in merged.columns else 0
        
        # Calculate hours worked using config durations
        trt_duration = VISIT_DURATIONS.get('trt', 20)
        hrt_duration = VISIT_DURATIONS.get('hrt', 20)
        other_duration = VISIT_DURATIONS.get('other', 20)
        merged['Hours Worked'] = ((merged['TRT Visits'] * trt_duration + merged['HRT Visits'] * hrt_duration + merged['Other Visits'] * other_duration) / 60).round(2)
        
        # Calculate percentage: (Hours Worked / Gusto Hours) * 100
        merged['Percentage Worked'] = ((merged['Hours Worked'] / merged['Total hours']) * 100).round(1)
        
        result = merged[['Name', 'Total hours', 'TRT Visits', 'HRT Visits', 'Total Visits', 'Hours Worked', 'Percentage Worked']].copy()
        result.columns = ['Provider', 'Gusto Hours', 'TRT Visits (20 min)', 'HRT Visits (20 min)', 'Total Visits', 'Hours Worked', 'Percentage Worked']
        result = result.sort_values('Provider', ascending=True)
    
    # Add N/A providers with calculated hours from visits
    if not visits.empty and 'Provider' in visits.columns:
        # Apply name mappings to visits before normalizing (if not already done)
        if 'Name_norm' not in visits.columns:
            def apply_visit_mapping(name):
                if pd.isna(name):
                    return name
                name_lower = str(name).lower().strip()
                return VISIT_NAME_MAPPINGS.get(name_lower, name)
            visits['Provider_mapped'] = visits['Provider'].apply(apply_visit_mapping)
            visits['Name_norm'] = visits['Provider_mapped'].apply(normalize_name)
        for provider_name in PROVIDERS_NA_HOURS:
            provider_norm = normalize_name(provider_name)
            # Check if there's an alternative name mapping for visits
            lookup_name = VISIT_NAME_MAPPINGS.get(provider_norm, provider_norm)
            if lookup_name != provider_norm:
                lookup_name = normalize_name(lookup_name)
            visit_match = visits[visits['Name_norm'] == lookup_name]
            if not visit_match.empty:
                row = visit_match.iloc[0]
                trt = row['TRT'] if 'TRT' in row else 0
                hrt = row['HRT'] if 'HRT' in row else 0
                other = row['Other'] if 'Other' in row else 0
                total = row['Total'] if 'Total' in row else (trt + hrt + other)
            else:
                # No visits found - show 0 hours worked
                trt, hrt, other, total = 0, 0, 0, 0
            
            hours_worked = round((trt * 20 + hrt * 20 + other * 20) / 60, 2)
            
            new_row = pd.DataFrame({
                'Provider': [provider_name],
                'Gusto Hours': ['N/A'],
                'TRT Visits (20 min)': [trt],
                'HRT Visits (20 min)': [hrt],
                'Total Visits': [total],
                'Hours Worked': [hours_worked],
                'Percentage Worked': ['N/A']
            })
            result = pd.concat([result, new_row], ignore_index=True)
    else:
        # If visits is empty, just add N/A providers with 0s
        for provider_name in PROVIDERS_NA_HOURS:
            new_row = pd.DataFrame({
                'Provider': [provider_name],
                'Gusto Hours': ['N/A'],
                'TRT Visits (20 min)': [0],
                'HRT Visits (20 min)': [0],
                'Total Visits': [0],
                'Hours Worked': [0],
                'Percentage Worked': ['N/A']
            })
            result = pd.concat([result, new_row], ignore_index=True)
    
    return result

def get_percentage_hours_worked(hours_worked_df):
    """Create a simplified sheet showing only percentage of hours worked."""
    if hours_worked_df.empty:
        return pd.DataFrame(columns=['Provider', 'Gusto Hours', 'Total Visits', 'Hours Worked (from visits)', 'Percentage Worked'])
    
    # Create simplified dataframe
    result = pd.DataFrame({
        'Provider': hours_worked_df['Provider'],
        'Gusto Hours': hours_worked_df['Gusto Hours'],
        'Total Visits': hours_worked_df['Total Visits'],
        'Hours Worked (from visits)': hours_worked_df['Hours Worked'],
        'Percentage Worked': hours_worked_df['Percentage Worked']
    })
    
    # Sort by provider name alphabetically
    result = result.sort_values('Provider', ascending=True)
    
    return result


def read_file_as_dataframe(file_obj, skiprows=0, diagnostic=None):
    """Read a file as DataFrame, handling both CSV and XLS formats."""
    filename = file_obj.filename if hasattr(file_obj, 'filename') else 'unknown'
    filename_lower = filename.lower()
    is_excel = filename_lower.endswith('.xls') or filename_lower.endswith('.xlsx')
    
    if diagnostic:
        diagnostic.add_info(f"Reading file: {filename}", {'is_excel': is_excel, 'skiprows': skiprows})
    
    file_obj.seek(0)
    
    if is_excel:
        # Try reading as Excel first
        try:
            df = pd.read_excel(file_obj, skiprows=skiprows)
            if diagnostic:
                diagnostic.add_info(f"Successfully read Excel file", {
                    'rows': len(df),
                    'columns': list(df.columns),
                    'file': filename
                })
            return df
        except Exception as e:
            # If that fails, it might be HTML disguised as XLS (OnceHub style)
            if diagnostic:
                diagnostic.add_info(f"Excel read failed, trying HTML parsing", {'error': str(e), 'file': filename})
            file_obj.seek(0)
            content = None
            encoding_used = None
            for encoding in ['utf-16', 'utf-8', 'latin-1', 'cp1252']:
                try:
                    file_obj.seek(0)
                    content = file_obj.read().decode(encoding)
                    encoding_used = encoding
                    break
                except (UnicodeDecodeError, UnicodeError):
                    continue
            if content:
                if diagnostic:
                    diagnostic.add_info(f"Decoded file as HTML with encoding: {encoding_used}", {'file': filename})
                # Try to read as HTML table
                try:
                    tables = pd.read_html(io.StringIO(content))
                    if tables:
                        df = tables[0]
                        if skiprows > 0:
                            df = df.iloc[skiprows:]
                            df.columns = df.iloc[0]
                            df = df.iloc[1:].reset_index(drop=True)
                        if diagnostic:
                            diagnostic.add_info(f"Successfully parsed HTML table", {
                                'rows': len(df),
                                'columns': list(df.columns),
                                'file': filename
                            })
                        return df
                except Exception as e2:
                    if diagnostic:
                        diagnostic.add_error(f"Could not parse HTML table", {'error': str(e2), 'file': filename})
                    pass
            if diagnostic:
                diagnostic.add_error(f"Could not read Excel file", {'file': filename, 'error': str(e)})
            raise ValueError("Could not read Excel file")
    else:
        # Read as CSV
        try:
            file_obj.seek(0)
            # Try different approaches for reading CSV with error handling
            df = None
            last_error = None
            
            # Method 1: Try with on_bad_lines='skip' and C engine (fastest)
            try:
                file_obj.seek(0)
                df = pd.read_csv(file_obj, skiprows=skiprows, on_bad_lines='skip')
            except (TypeError, ValueError, AssertionError) as e1:
                last_error = e1
                # Method 2: Try with python engine
                try:
                    file_obj.seek(0)
                    df = pd.read_csv(file_obj, skiprows=skiprows, on_bad_lines='skip', engine='python')
                except (TypeError, ValueError, AssertionError) as e2:
                    last_error = e2
                    # Method 3: Try with error_bad_lines (older pandas)
                    try:
                        file_obj.seek(0)
                        df = pd.read_csv(file_obj, skiprows=skiprows, error_bad_lines=False, warn_bad_lines=False)
                    except (TypeError, ValueError, AssertionError) as e3:
                        last_error = e3
                        # Method 4: Try with python engine and error_bad_lines
                        try:
                            file_obj.seek(0)
                            df = pd.read_csv(file_obj, skiprows=skiprows, error_bad_lines=False, warn_bad_lines=False, engine='python')
                        except Exception as e4:
                            last_error = e4
                            # Method 5: Last resort - try without skiprows first, then manually skip
                            try:
                                file_obj.seek(0)
                                df = pd.read_csv(file_obj, on_bad_lines='skip')
                                if skiprows > 0:
                                    df = df.iloc[skiprows:].reset_index(drop=True)
                            except Exception as e5:
                                last_error = e5
                                raise
            
            if df is None or df.empty:
                raise ValueError("CSV file appears to be empty after reading")
                
            if diagnostic:
                diagnostic.add_info(f"Successfully read CSV file", {
                    'rows': len(df),
                    'columns': list(df.columns),
                    'file': filename
                })
            return df
        except Exception as e:
            error_msg = str(e) if str(e) else f"Unknown error: {type(e).__name__}"
            if diagnostic:
                diagnostic.add_error(f"Could not read CSV file", {'error': error_msg, 'file': filename})
            raise ValueError(f"Could not read CSV file: {error_msg}")


def normalize_name_for_fuzzy(name):
    """Normalize a name for fuzzy matching comparison"""
    if pd.isna(name):
        return ''
    name = str(name).strip()
    name = re.sub(r'\s+(NP|FNP-C|MD|PA|LLC|Inc\.?|INC\.?|PLLC)$', '', name, flags=re.IGNORECASE)
    name = re.sub(r',\s*NP$', '', name, flags=re.IGNORECASE)
    return name.lower().strip()


def find_best_fuzzy_match(source_name, target_list, threshold=85):
    """
    Find the best fuzzy match for a name in a list of target names.
    Returns (matched_name, confidence_score) or (None, 0) if no good match.
    """
    if not FUZZY_AVAILABLE or not source_name or not target_list:
        return None, 0
    
    # Use token_sort_ratio which handles word order differences
    best_match = process.extractOne(
        source_name, 
        target_list, 
        scorer=fuzz.token_sort_ratio
    )
    
    if best_match and best_match[1] >= threshold:
        return best_match[0], best_match[1]
    
    return None, 0


def generate_name_mapping_suggestions(source_names, target_names, diagnostic=None):
    """
    Generate suggested name mappings using fuzzy matching.
    Returns: (auto_mappings, manual_review_suggestions)
    """
    
    auto_mappings = {}
    manual_suggestions = []
    
    for source_name in source_names:
        # Normalize for comparison
        normalized_source = normalize_name_for_fuzzy(source_name)
        
        # Try to find exact match first
        exact_match = None
        for target in target_names:
            if normalize_name_for_fuzzy(target) == normalized_source:
                exact_match = target
                break
        
        if exact_match:
            # Exact match found (after normalization)
            if source_name.lower() != exact_match.lower():
                # Names are different but normalize to same thing
                auto_mappings[source_name] = exact_match
                if diagnostic:
                    diagnostic.add_info(f"Exact match (normalized): {source_name} → {exact_match}")
            continue
        
        # No exact match, try fuzzy matching
        best_match, confidence = find_best_fuzzy_match(source_name, target_names)
        
        if best_match:
            if confidence >= 95:
                # High confidence - auto-apply
                auto_mappings[source_name] = best_match
                if diagnostic:
                    diagnostic.add_info(f"Auto-matched (high confidence): {source_name} → {best_match}", 
                                      {'confidence': confidence})
            elif confidence >= 85:
                # Medium confidence - suggest for review
                manual_suggestions.append({
                    'Source Name': source_name,
                    'Suggested Match': best_match,
                    'Confidence': f"{confidence}%",
                    'Status': 'Review Recommended',
                    'Action': 'Add to config.yaml if correct'
                })
                if diagnostic:
                    diagnostic.add_warning(f"Possible match needs review: {source_name} → {best_match}", 
                                         {'confidence': confidence})
    
    suggestions_df = pd.DataFrame(manual_suggestions) if manual_suggestions else pd.DataFrame(
        columns=['Source Name', 'Suggested Match', 'Confidence', 'Status', 'Action']
    )
    
    if diagnostic:
        diagnostic.add_stat('fuzzy_auto_matches', len(auto_mappings))
        diagnostic.add_stat('fuzzy_manual_review', len(manual_suggestions))
    
    return auto_mappings, suggestions_df


def apply_fuzzy_matching_to_providers(gusto_names, visit_names, diagnostic=None):
    """
    Apply fuzzy matching to match Gusto providers with visit providers.
    Updates GUSTO_NAME_MAPPINGS with high-confidence matches.
    Returns suggestions DataFrame for manual review.
    """
    
    if not FUZZY_AVAILABLE:
        if diagnostic:
            diagnostic.add_info("Fuzzy matching not available (fuzzywuzzy not installed)")
        return pd.DataFrame(columns=['Source Name', 'Suggested Match', 'Confidence', 'Status', 'Action'])
    
    if diagnostic:
        diagnostic.add_info("Running fuzzy name matching", {
            'gusto_providers': len(gusto_names),
            'visit_providers': len(visit_names)
        })
    
    # Get existing mapped names to avoid duplicates
    already_mapped = set(GUSTO_NAME_MAPPINGS.keys())
    unmapped_gusto = [name for name in gusto_names if name.lower() not in already_mapped]
    
    # Generate suggestions
    auto_mappings, suggestions_df = generate_name_mapping_suggestions(
        unmapped_gusto, 
        visit_names, 
        diagnostic
    )
    
    # Apply auto mappings
    for source, target in auto_mappings.items():
        GUSTO_NAME_MAPPINGS[source.lower()] = target
        config['gusto_mappings'][source.lower()] = target
        if diagnostic:
            diagnostic.add_info(f"Applied automatic mapping: {source} → {target}")
    
    return suggestions_df


def run_quality_checks(data_dict, diagnostic=None):
    """
    Run comprehensive data quality checks on the report data.
    Returns a DataFrame of quality issues found.
    """
    
    if diagnostic:
        diagnostic.add_info("Running data quality checks")
    
    issues = []
    
    # Extract data
    hours_df = data_dict.get('hours_worked', pd.DataFrame())
    visits_df = data_dict.get('doxy_visits', pd.DataFrame())
    visits_program = data_dict.get('visits_by_program', pd.DataFrame())
    gusto_df = data_dict.get('gusto_hours', pd.DataFrame())
    performance_df = data_dict.get('performance_metrics', pd.DataFrame())
    
    # Get thresholds from config
    thresholds = QUALITY_THRESHOLDS
    
    # CHECK 1: Percentage worked > threshold
    if not hours_df.empty and 'Percentage Worked' in hours_df.columns:
        max_threshold = thresholds.get('max_percentage_worked', 110)
        over_hours = hours_df[
            (hours_df['Percentage Worked'] != 'N/A') & 
            (pd.to_numeric(hours_df['Percentage Worked'], errors='coerce') > max_threshold)
        ]
        
        if len(over_hours) > 0:
            issues.append({
                'Severity': '⚠ WARNING',
                'Check': 'Hours Worked Validation',
                'Issue': f"{len(over_hours)} provider(s) worked >{max_threshold}% of Gusto hours",
                'Affected Items': ', '.join(over_hours['Provider'].tolist()),
                'Recommended Action': 'Verify Gusto hours cover correct date range. May indicate part-time week or data mismatch.'
            })
            if diagnostic:
                diagnostic.add_warning(f"{len(over_hours)} providers exceeded hours threshold", {
                    'providers': over_hours['Provider'].tolist(),
                    'percentages': over_hours['Percentage Worked'].tolist()
                })
    
    # CHECK 2: TRT/HRT ratio validation
    if not visits_program.empty and 'TRT' in visits_program.columns and 'HRT' in visits_program.columns:
        total_trt = visits_program['TRT'].sum()
        total_hrt = visits_program['HRT'].sum()
        total_both = total_trt + total_hrt
        
        if total_both > 0:
            trt_ratio = total_trt / total_both
            min_ratio = thresholds.get('trt_ratio_min', 0.40)
            max_ratio = thresholds.get('trt_ratio_max', 0.85)
            
            if trt_ratio < min_ratio or trt_ratio > max_ratio:
                issues.append({
                    'Severity': '⚠ WARNING',
                    'Check': 'TRT/HRT Ratio',
                    'Issue': f"TRT visits are {trt_ratio:.1%} of total (expected {min_ratio:.0%}-{max_ratio:.0%})",
                    'Affected Items': f"TRT: {total_trt}, HRT: {total_hrt}",
                    'Recommended Action': 'Verify Event Type categorization is working correctly. Check Account Detail export.'
                })
                if diagnostic:
                    diagnostic.add_warning("TRT/HRT ratio outside expected range", {
                        'trt_ratio': f"{trt_ratio:.1%}",
                        'trt_count': total_trt,
                        'hrt_count': total_hrt
                    })
    
    # CHECK 3: Providers with Gusto hours but no visits
    if not hours_df.empty and 'Total Visits' in hours_df.columns and 'Gusto Hours' in hours_df.columns:
        zero_visits = hours_df[
            (hours_df['Total Visits'] == 0) & 
            (hours_df['Gusto Hours'] != 'N/A') &
            (pd.to_numeric(hours_df['Gusto Hours'], errors='coerce') > 0)
        ]
        
        if len(zero_visits) > 0:
            issues.append({
                'Severity': 'ℹ INFO',
                'Check': 'Visit Coverage',
                'Issue': f"{len(zero_visits)} provider(s) have Gusto hours but no visits",
                'Affected Items': ', '.join(zero_visits['Provider'].tolist()),
                'Recommended Action': 'Normal if provider was on administrative duties. Otherwise verify visit data.'
            })
            if diagnostic:
                diagnostic.add_info(f"{len(zero_visits)} providers with hours but no visits", {
                    'providers': zero_visits['Provider'].tolist()
                })
    
    # CHECK 4: Providers with visits but no Gusto hours
    if not visits_df.empty and not gusto_df.empty:
        visit_providers = set(visits_df['Provider name'].dropna())
        gusto_providers = set(gusto_df['Name'].dropna())
        
        missing_gusto = visit_providers - gusto_providers
        # Filter out providers that are intentionally in NA list
        missing_gusto = [p for p in missing_gusto if p not in PROVIDERS_NA_HOURS]
        
        if missing_gusto:
            issues.append({
                'Severity': '🔴 ERROR',
                'Check': 'Provider Coverage',
                'Issue': f"{len(missing_gusto)} provider(s) have visits but no Gusto record",
                'Affected Items': ', '.join(sorted(missing_gusto)),
                'Recommended Action': 'Add to Gusto system OR add to config.yaml na_providers list if intentional.'
            })
            if diagnostic:
                diagnostic.add_error("Providers missing from Gusto", {
                    'count': len(missing_gusto),
                    'providers': list(missing_gusto)
                })
    
    # CHECK 5: Average duration validation
    if not performance_df.empty and 'Avg Duration (min)' in performance_df.columns:
        min_duration = thresholds.get('min_avg_duration', 5)
        max_duration = thresholds.get('max_avg_duration', 60)
        
        low_duration = performance_df[performance_df['Avg Duration (min)'] < min_duration]
        high_duration = performance_df[performance_df['Avg Duration (min)'] > max_duration]
        
        if len(low_duration) > 0:
            issues.append({
                'Severity': '⚠ WARNING',
                'Check': 'Visit Duration - Low',
                'Issue': f"{len(low_duration)} provider(s) have average duration <{min_duration} min",
                'Affected Items': ', '.join(low_duration['Provider'].tolist()),
                'Recommended Action': 'May indicate incomplete visits or technical issues. Review Doxy data.'
            })
        
        if len(high_duration) > 0:
            issues.append({
                'Severity': '⚠ WARNING',
                'Check': 'Visit Duration - High',
                'Issue': f"{len(high_duration)} provider(s) have average duration >{max_duration} min",
                'Affected Items': ', '.join(high_duration['Provider'].tolist()),
                'Recommended Action': 'Verify duration data is accurate. May be valid for complex cases.'
            })
    
    # CHECK 6: Unusually low visit counts
    if not visits_df.empty and len(visits_df) > 3:
        avg_visits = visits_df['Total Visits'].mean()
        min_ratio = thresholds.get('min_visits_ratio', 0.3)
        threshold_visits = avg_visits * min_ratio
        
        low_visits = visits_df[visits_df['Total Visits'] < threshold_visits]
        
        if len(low_visits) > 3:  # Only flag if multiple providers affected
            issues.append({
                'Severity': 'ℹ INFO',
                'Check': 'Visit Count Anomaly',
                'Issue': f"{len(low_visits)} provider(s) have unusually low visits (<{threshold_visits:.0f}, avg is {avg_visits:.0f})",
                'Affected Items': ', '.join(low_visits['Provider name'].tolist()),
                'Recommended Action': 'Verify these providers were working full schedules. May be normal for part-time.'
            })
    
    # CHECK 7: Data source consistency
    if not visits_df.empty and not visits_program.empty:
        doxy_total = visits_df['Total Visits'].sum()
        account_total = visits_program['Total'].sum() if 'Total' in visits_program.columns else 0
        
        if account_total > 0:
            difference_pct = abs(doxy_total - account_total) / max(doxy_total, account_total) * 100
            
            if difference_pct > 10:  # More than 10% difference
                issues.append({
                    'Severity': '⚠ WARNING',
                    'Check': 'Data Source Consistency',
                    'Issue': f"Doxy visits ({doxy_total}) and Account Detail visits ({account_total}) differ by {difference_pct:.1f}%",
                    'Affected Items': f"Difference: {abs(doxy_total - account_total)} visits",
                    'Recommended Action': 'Verify both sources cover the same date range and providers. Check for missing data.'
                })
                if diagnostic:
                    diagnostic.add_warning("Visit count mismatch between sources", {
                        'doxy_total': doxy_total,
                        'account_total': account_total,
                        'difference_pct': f"{difference_pct:.1f}%"
                    })
    
    # Create DataFrame from issues
    if issues:
        quality_df = pd.DataFrame(issues)
        if diagnostic:
            error_count = len(quality_df[quality_df['Severity'].str.contains('ERROR', na=False)])
            warning_count = len(quality_df[quality_df['Severity'].str.contains('WARNING', na=False)])
            
            diagnostic.add_stat('quality_errors', error_count)
            diagnostic.add_stat('quality_warnings', warning_count)
            diagnostic.add_info(f"Quality checks complete: {error_count} errors, {warning_count} warnings")
    else:
        quality_df = pd.DataFrame(columns=[
            'Severity', 'Check', 'Issue', 'Affected Items', 'Recommended Action'
        ])
        if diagnostic:
            diagnostic.add_info("Quality checks complete: No issues found")
    
    return quality_df


def create_provider_reconciliation(doxy_df, oncehub_df, visits_program_df, gusto_df, diagnostic=None):
    """
    Create provider reconciliation report showing which providers appear in which sources.
    Returns a DataFrame showing provider presence across all systems.
    """
    
    if diagnostic:
        diagnostic.add_info("Creating provider reconciliation report")
    
    # Extract unique providers from each source
    doxy_providers = set()
    if doxy_df is not None and not doxy_df.empty:
        if 'Provider name' in doxy_df.columns:
            doxy_providers = set(doxy_df['Provider name'].dropna().unique())
        elif 'Provider' in doxy_df.columns:
            doxy_providers = set(doxy_df['Provider'].dropna().unique())
    
    oncehub_providers = set()
    if oncehub_df is not None and not oncehub_df.empty and 'Provider' in oncehub_df.columns:
        oncehub_providers = set(oncehub_df['Provider'].dropna().unique())
    
    account_providers = set()
    if visits_program_df is not None and not visits_program_df.empty and 'Provider' in visits_program_df.columns:
        account_providers = set(visits_program_df['Provider'].dropna().unique())
    
    gusto_providers = set()
    if gusto_df is not None and not gusto_df.empty and 'Name' in gusto_df.columns:
        gusto_providers = set(gusto_df['Name'].dropna().unique())
    
    # Combine all unique providers
    all_providers = doxy_providers | oncehub_providers | account_providers | gusto_providers
    
    if diagnostic:
        diagnostic.add_info("Provider reconciliation counts", {
            'doxy': len(doxy_providers),
            'oncehub': len(oncehub_providers),
            'account': len(account_providers),
            'gusto': len(gusto_providers),
            'total_unique': len(all_providers)
        })
    
    reconciliation_data = []
    
    for provider in sorted(all_providers):
        # Check presence in each system
        in_doxy = provider in doxy_providers
        in_oncehub = provider in oncehub_providers
        in_account = provider in account_providers
        in_gusto = provider in gusto_providers
        
        # Determine status and required action
        status = ""
        action = ""
        severity = "INFO"
        
        if in_doxy and in_oncehub and in_account and in_gusto:
            status = "✓ Complete - All Systems"
            action = ""
            severity = "OK"
        elif not in_gusto and (in_doxy or in_oncehub):
            status = "⚠ Missing Gusto Hours"
            action = "Add to Gusto or add to config.yaml na_providers list"
            severity = "WARNING"
        elif in_gusto and not (in_doxy or in_oncehub):
            status = "⚠ No Visits This Week"
            action = "Normal if provider didn't work"
            severity = "INFO"
        elif not in_account and (in_doxy or in_oncehub):
            status = "⚠ Missing from Account Detail"
            action = "Check Account Detail export or add name mapping"
            severity = "WARNING"
        elif in_account and not (in_doxy or in_oncehub):
            status = "⚠ In Account but No Visit System"
            action = "Verify visit systems or check name variations"
            severity = "WARNING"
        else:
            status = "⚠ Partial Coverage"
            action = "Review presence in each system"
            severity = "WARNING"
        
        reconciliation_data.append({
            'Provider': provider,
            'In Doxy': '✓' if in_doxy else '✗',
            'In OnceHub': '✓' if in_oncehub else '✗',
            'In Account Detail': '✓' if in_account else '✗',
            'In Gusto': '✓' if in_gusto else '✗',
            'Status': status,
            'Recommended Action': action,
            'Severity': severity
        })
    
    reconciliation_df = pd.DataFrame(reconciliation_data)
    
    # Add summary statistics
    if diagnostic:
        complete_count = len(reconciliation_df[reconciliation_df['Severity'] == 'OK'])
        warning_count = len(reconciliation_df[reconciliation_df['Severity'] == 'WARNING'])
        
        diagnostic.add_stat('reconciliation_complete', complete_count)
        diagnostic.add_stat('reconciliation_warnings', warning_count)
        
        if warning_count > 0:
            diagnostic.add_warning(f"{warning_count} providers have incomplete coverage across systems")
    
    return reconciliation_df


def generate_report(doxy_file, account_file, gusto_file, booking_file):
    """Generate the complete Excel report with detailed error handling."""
    diagnostic = DiagnosticLog()
    diagnostic.add_info("Starting report generation")
    
    errors = []
    
    # Read Doxy Report (CSV or XLS)
    try:
        doxy_df = read_file_as_dataframe(doxy_file, diagnostic=diagnostic)
        if 'Provider name' not in doxy_df.columns:
            error_msg = "Doxy Report missing 'Provider name' column"
            errors.append(error_msg)
            diagnostic.add_error(error_msg, {'available_columns': list(doxy_df.columns)})
        if 'Duration' not in doxy_df.columns:
            error_msg = "Doxy Report missing 'Duration' column"
            errors.append(error_msg)
            diagnostic.add_error(error_msg, {'available_columns': list(doxy_df.columns)})
    except Exception as e:
        error_msg = f"Error reading Doxy Report: {str(e)}"
        errors.append(error_msg)
        diagnostic.add_error(error_msg)
        doxy_df = None
    
    # Read Account Detail Report (try different encodings or as Excel)
    account_content = None
    account_is_csv = account_file.filename.lower().endswith('.csv')
    account_df = None
    
    if account_is_csv:
        # Try to read as CSV with different encodings
        for encoding in ['utf-8', 'utf-16', 'latin-1', 'cp1252']:
            try:
                account_file.seek(0)
                raw_content = account_file.read()
                account_content = raw_content.decode(encoding)
                # Verify the content is not empty and has actual data
                if account_content and len(account_content.strip()) > 100:
                    if diagnostic:
                        diagnostic.add_info(f"Successfully decoded Account Detail Report with {encoding}", {
                            'content_length': len(account_content),
                            'first_100_chars': account_content[:100]
                        })
                    break
                else:
                    account_content = None
            except (UnicodeDecodeError, UnicodeError):
                continue
        
        if account_content is None:
            error_msg = "Could not decode Account Detail Report - try a different file format"
            errors.append(error_msg)
            diagnostic.add_error(error_msg)
    else:
        # Try to read as Excel file first
        try:
            account_file.seek(0)
            # Check if it's actually HTML by trying to decode and look for HTML tags
            is_html = False
            for encoding in ['utf-16', 'utf-8', 'latin-1', 'cp1252']:
                try:
                    account_file.seek(0)
                    peek_content = account_file.read(500).decode(encoding, errors='ignore')
                    if '<table' in peek_content.lower() or '<html' in peek_content.lower():
                        is_html = True
                        # Read full content
                        account_file.seek(0)
                        account_content = account_file.read().decode(encoding)
                        account_is_csv = False  # Use HTML parsing
                        break
                except (UnicodeDecodeError, UnicodeError):
                    continue
            
            if not is_html:
                # Real Excel file
                account_file.seek(0)
                account_df = read_file_as_dataframe(account_file, diagnostic=diagnostic)
                # Convert to string format for get_visits_by_program
                account_content = account_df.to_csv(index=False)
                account_is_csv = True  # Treat as CSV format for processing
        except Exception as e:
            # If Excel read fails, try reading as HTML/text
            try:
                account_file.seek(0)
                encoding_used = None
                for encoding in ['utf-16', 'utf-8', 'latin-1', 'cp1252']:
                    try:
                        account_file.seek(0)
                        account_content = account_file.read().decode(encoding)
                        encoding_used = encoding
                        account_is_csv = False  # Use HTML parsing
                        break
                    except (UnicodeDecodeError, UnicodeError):
                        continue
                if encoding_used and diagnostic:
                    diagnostic.add_info(f"Decoded Account Detail Report with encoding: {encoding_used}")
            except Exception as e2:
                error_msg = f"Could not read Account Detail Report: {str(e)}"
                errors.append(error_msg)
                diagnostic.add_error(error_msg, {'secondary_error': str(e2)})
    
    # Read Gusto Hours (CSV or XLS, skip header rows)
    try:
        gusto_df = read_file_as_dataframe(gusto_file, skiprows=8, diagnostic=diagnostic)
    except Exception as e:
        error_msg = f"Error reading Gusto file: {str(e)}"
        errors.append(error_msg)
        diagnostic.add_error(error_msg)
        gusto_df = None
    
    # Read Booking Summary (OnceHub) - CSV or XLS (OPTIONAL)
    booking_df = None
    if booking_file and booking_file.filename:
        try:
            booking_df = read_file_as_dataframe(booking_file, diagnostic=diagnostic)
            # Try to find the booking page column with flexible matching
            booking_col = None
            for col in booking_df.columns:
                col_lower = str(col).lower()
                if 'booking' in col_lower or 'page' in col_lower or 'provider' in col_lower or 'name' in col_lower:
                    booking_col = col
                    break
            
            if booking_col and booking_col != 'Booking page':
                booking_df = booking_df.rename(columns={booking_col: 'Booking page'})
            elif 'Booking page' not in booking_df.columns:
                # Use first column as provider name if no match found
                booking_df = booking_df.rename(columns={booking_df.columns[0]: 'Booking page'})
        except Exception as e:
            # OnceHub is optional - just log and continue
            booking_df = None
    
    # If there are critical errors, raise them
    if errors:
        raise ValueError("\n".join(errors))
    
    # Generate all sections
    doxy_visits = get_doxy_visits(doxy_df, diagnostic=diagnostic)
    doxy_providers = doxy_visits['Provider name'].tolist() if not doxy_visits.empty else []
    
    # OnceHub is optional
    oncehub_visits = get_oncehub_visits(booking_df, diagnostic=diagnostic) if booking_df is not None else None
    
    # Extract date range from Doxy report for filtering Account Detail Report
    # Use Doxy data to determine date range (all days included)
    start_date = None
    end_date = None
    if doxy_df is not None and 'Date' in doxy_df.columns:
        try:
            # Process Doxy data to get the correct date range
            filtered_doxy = filter_out_upload_day(doxy_df.copy())
            if not filtered_doxy.empty and 'Date' in filtered_doxy.columns:
                doxy_dates = pd.to_datetime(filtered_doxy['Date'], errors='coerce')
                doxy_dates = doxy_dates[doxy_dates.notna()]
                if len(doxy_dates) > 0:
                    start_date = doxy_dates.min()
                    end_date = doxy_dates.max()
        except:
            pass
    
    # Get visits by program - handle None account_content
    if account_content is not None:
        visits_by_program = get_visits_by_program(account_content, is_csv=account_is_csv, start_date=start_date, end_date=end_date, diagnostic=diagnostic)
    else:
        visits_by_program = pd.DataFrame(columns=['Provider', 'TRT', 'HRT', 'Other', 'Total'])
        error_msg = "Account Detail Report could not be parsed - Visits by Program will be empty"
        errors.append(error_msg)
        diagnostic.add_error(error_msg)
    
    gusto_hours = get_gusto_hours(gusto_df, doxy_providers, diagnostic=diagnostic)
    performance_metrics = get_doxy_performance_metrics(doxy_df, diagnostic=diagnostic)
    
    # Run fuzzy matching (may update gusto_hours mappings)
    gusto_provider_names = gusto_hours['Name'].unique().tolist() if not gusto_hours.empty else []
    visit_provider_names = visits_by_program['Provider'].unique().tolist() if not visits_by_program.empty else []
    
    fuzzy_suggestions = apply_fuzzy_matching_to_providers(
        gusto_provider_names,
        visit_provider_names,
        diagnostic
    )
    
    # If we found auto-matches, re-run the hours calculation with updated mappings
    if len(GUSTO_NAME_MAPPINGS) > len(config.get('gusto_mappings', {})):
        diagnostic.add_info("Re-calculating with fuzzy-matched names")
        gusto_hours = get_gusto_hours(gusto_df, doxy_providers, diagnostic=diagnostic)
    
    hours_worked = get_hours_worked(gusto_hours, visits_by_program, diagnostic=diagnostic)
    percentage_hours = get_percentage_hours_worked(hours_worked)
    
    # Create provider reconciliation report
    reconciliation = create_provider_reconciliation(
        doxy_visits, 
        oncehub_visits, 
        visits_by_program, 
        gusto_hours,
        diagnostic
    )
    
    # Run quality checks on all data
    quality_alerts = run_quality_checks({
        'hours_worked': hours_worked,
        'doxy_visits': doxy_visits,
        'visits_by_program': visits_by_program,
        'gusto_hours': gusto_hours,
        'performance_metrics': performance_metrics
    }, diagnostic)
    
    # Calculate stats for response
    stats = {
        'providers': len(doxy_visits),
        'total_visits': int(doxy_visits['Total Visits'].sum()),
        'sheets': 7
    }
    
    # Create Excel file in memory
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. Provider Reconciliation (most important for gaps)
        reconciliation.to_excel(writer, sheet_name='Provider Reconciliation', index=False)
        
        # 2. Quality Alerts (critical issues)
        if not quality_alerts.empty:
            quality_alerts.to_excel(writer, sheet_name='⚠️ Quality Alerts', index=False)
        
        # 3. Suggested Name Mappings (from fuzzy matching)
        if not fuzzy_suggestions.empty:
            fuzzy_suggestions.to_excel(writer, sheet_name='Suggested Name Mappings', index=False)
        
        # 4-10. Regular data sheets
        doxy_visits.to_excel(writer, sheet_name='Doxy Visits', index=False)
        if oncehub_visits is not None:
            oncehub_visits.to_excel(writer, sheet_name='OnceHub Visits', index=False)
        visits_by_program.to_excel(writer, sheet_name='Visits by Program', index=False)
        gusto_hours.to_excel(writer, sheet_name='Gusto Hours', index=False)
        performance_metrics.to_excel(writer, sheet_name='Doxy Performance Metrics', index=False)
        hours_worked.to_excel(writer, sheet_name='Hours Worked', index=False)
        percentage_hours.to_excel(writer, sheet_name='% of hours worked', index=False)
        
        # 11. Diagnostics (last, for debugging)
        diag_summary = diagnostic.get_summary()
        diag_data = []
        
        # Add errors
        for err in diag_summary['errors']:
            diag_data.append({
                'Type': 'ERROR',
                'Message': err['message'],
                'Context': str(err.get('context', ''))
            })
        
        # Add warnings
        for warn in diag_summary['warnings']:
            diag_data.append({
                'Type': 'WARNING',
                'Message': warn['message'],
                'Context': str(warn.get('context', ''))
            })
        
        # Add info (limit to most recent 20)
        for info in diag_summary['info'][-20:]:
            diag_data.append({
                'Type': 'INFO',
                'Message': info['message'],
                'Context': str(info.get('context', ''))
            })
        
        # Add stats
        for key, value in diag_summary['stats'].items():
            diag_data.append({
                'Type': 'STAT',
                'Message': key,
                'Context': str(value)
            })
        
        if diag_data:
            diag_df = pd.DataFrame(diag_data)
            diag_df.to_excel(writer, sheet_name='Diagnostics', index=False)
    
    output.seek(0)
    
    # Save report to reports directory (only if not in serverless environment)
    is_serverless = os.environ.get('VERCEL') == '1' or os.environ.get('AWS_LAMBDA_FUNCTION_NAME') or not os.path.exists('/tmp')
    
    if not is_serverless:
        reports_dir = 'reports'
        try:
            if not os.path.exists(reports_dir):
                os.makedirs(reports_dir)
            
            generation_date = datetime.now().strftime("%Y-%m-%d")
            report_filename = f'Weekly_Report_{generation_date}.xlsx'
            report_path = os.path.join(reports_dir, report_filename)
            
            # Save the report
            with open(report_path, 'wb') as f:
                f.write(output.getvalue())
            logger.info(f"Saved report to {report_path}")
        except Exception as e:
            logger.warning(f"Could not save report to disk (likely serverless): {e}")
    
    # Reset output for download
    output.seek(0)
    
    # Prepare preview data
    preview_data = {
        'doxy_visits': doxy_visits,
        'oncehub_visits': oncehub_visits,
        'visits_by_program': visits_by_program,
        'gusto_hours': gusto_hours,
        'performance_metrics': performance_metrics,
        'hours_worked': hours_worked,
        'percentage_hours': percentage_hours,
        'reconciliation': reconciliation,
        'quality_alerts': quality_alerts,
        'fuzzy_suggestions': fuzzy_suggestions,
        'diagnostics': diagnostic.get_summary()
    }
    
    return output, stats, preview_data


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Validate required files are present (booking_file is optional)
        required_files = ['doxy_file', 'account_file', 'gusto_file']
        optional_files = ['booking_file']
        missing_files = []
        
        for file_name in required_files:
            if file_name not in request.files or request.files[file_name].filename == '':
                missing_files.append(FILE_CONFIGS[file_name]['name'])
        
        if missing_files:
            flash(f"Missing required files: {', '.join(missing_files)}", 'error')
            return redirect(request.url)
        
        # Validate each required file
        validation_errors = []
        for file_name in required_files:
            file_obj = request.files[file_name]
            errors = validate_file(file_obj, FILE_CONFIGS[file_name])
            for error in errors:
                validation_errors.append(f"{FILE_CONFIGS[file_name]['name']}: {error}")
        
        # Validate optional files only if provided
        for file_name in optional_files:
            if file_name in request.files and request.files[file_name].filename != '':
                file_obj = request.files[file_name]
                errors = validate_file(file_obj, FILE_CONFIGS[file_name])
                for error in errors:
                    validation_errors.append(f"{FILE_CONFIGS[file_name]['name']}: {error}")
        
        if validation_errors:
            for error in validation_errors:
                flash(error, 'error')
            return redirect(request.url)
        
        doxy_file = request.files['doxy_file']
        account_file = request.files['account_file']
        gusto_file = request.files['gusto_file']
        booking_file = request.files.get('booking_file') if 'booking_file' in request.files and request.files['booking_file'].filename else None
        
        try:
            # Generate the report
            output, stats, preview_data = generate_report(doxy_file, account_file, gusto_file, booking_file)
            
            # Display diagnostics warnings/errors
            diag = preview_data.get('diagnostics', {})
            if diag.get('errors'):
                for err in diag['errors']:
                    flash(f"ERROR: {err['message']}", 'error')
            if diag.get('warnings'):
                for warn in diag['warnings'][:5]:  # Show first 5 warnings
                    flash(f"Warning: {warn['message']}", 'warning')
            
            # Always use 'Weekly Report (date of generation)'
            generation_date = datetime.now().strftime("%m-%d-%Y")
            report_name = f'Weekly Report ({generation_date})'
            
            # Convert preview data to JSON-serializable format
            def dataframe_to_dict(df):
                if df is None or df.empty:
                    return {'columns': [], 'data': []}
                return {
                    'columns': list(df.columns),
                    'data': df.fillna('').to_dict('records')
                }
            
            preview_json = {
                'summary': {
                    'total_providers': stats['providers'],
                    'total_doxy_visits': stats['total_visits'],
                    'total_hours_worked': round(preview_data['hours_worked']['Hours Worked'].sum(), 1) if not preview_data['hours_worked'].empty else 0
                },
                'sheets': [
                    {
                        'name': 'Doxy Visits',
                        'data': dataframe_to_dict(preview_data['doxy_visits'])
                    },
                    {
                        'name': 'OnceHub Visits',
                        'data': dataframe_to_dict(preview_data['oncehub_visits']),
                        'available': preview_data['oncehub_visits'] is not None
                    },
                    {
                        'name': 'Visits by Program',
                        'data': dataframe_to_dict(preview_data['visits_by_program'])
                    },
                    {
                        'name': 'Gusto Hours',
                        'data': dataframe_to_dict(preview_data['gusto_hours'])
                    },
                    {
                        'name': 'Doxy Performance Metrics',
                        'data': dataframe_to_dict(preview_data['performance_metrics'])
                    },
                    {
                        'name': 'Hours Worked',
                        'data': dataframe_to_dict(preview_data['hours_worked'])
                    },
                    {
                        'name': '% of hours worked',
                        'data': dataframe_to_dict(preview_data['percentage_hours'])
                    }
                ]
            }
            
            # Store preview data in session for display after download
            from flask import session
            session['last_preview'] = preview_json
            session['report_name'] = report_name
            
            return Response(
                output.getvalue(),
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                headers={
                    'Content-Disposition': f'attachment; filename="{report_name}.xlsx"',
                    'X-Report-Providers': str(stats['providers']),
                    'X-Report-Visits': str(stats['total_visits']),
                    'X-Has-Preview': 'true'
                }
            )
        except ValueError as e:
            # Specific validation errors
            for line in str(e).split('\n'):
                flash(line, 'error')
            return redirect(request.url)
        except Exception as e:
            flash(f'Unexpected error: {str(e)}', 'error')
            return redirect(request.url)
    
    return render_template('index.html')


@app.route('/get-preview', methods=['GET'])
def get_preview():
    """API endpoint to get the last generated report preview."""
    from flask import session
    if 'last_preview' in session:
        return jsonify(session['last_preview'])
    return jsonify({'error': 'No preview available'}), 404


@app.route('/api/config', methods=['GET'])
def get_config_endpoint():
    """Get current configuration"""
    return jsonify(config)


@app.route('/api/config', methods=['POST'])
def update_config_endpoint():
    """Update configuration"""
    try:
        new_config = request.json
        
        # Validate config structure
        required_keys = ['gusto_mappings', 'visit_mappings', 'na_providers', 'excluded_names']
        for key in required_keys:
            if key not in new_config:
                return jsonify({'error': f'Missing required key: {key}'}), 400
        
        # Save to file
        if save_config(new_config):
            # Reload config
            global config, GUSTO_NAME_MAPPINGS, VISIT_NAME_MAPPINGS, PROVIDERS_NA_HOURS, EXCLUDED_NAMES, QUALITY_THRESHOLDS, VISIT_DURATIONS
            config = new_config
            GUSTO_NAME_MAPPINGS = config.get('gusto_mappings', {})
            VISIT_NAME_MAPPINGS = config.get('visit_mappings', {})
            PROVIDERS_NA_HOURS = config.get('na_providers', [])
            EXCLUDED_NAMES = config.get('excluded_names', [])
            QUALITY_THRESHOLDS = config.get('quality_thresholds', {})
            VISIT_DURATIONS = config.get('visit_durations', {'trt': 20, 'hrt': 20, 'other': 20})
            
            return jsonify({'success': True, 'message': 'Configuration updated'})
        else:
            return jsonify({'error': 'Failed to save configuration'}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/config/add-mapping', methods=['POST'])
def add_mapping():
    """Add a single name mapping"""
    try:
        data = request.json
        mapping_type = data.get('type')  # 'gusto' or 'visits'
        source = data.get('source')
        target = data.get('target')
        
        if not mapping_type or not source or not target:
            return jsonify({'error': 'Missing required fields: type, source, target'}), 400
        
        if mapping_type == 'gusto':
            GUSTO_NAME_MAPPINGS[source.lower().strip()] = target.strip()
            config['gusto_mappings'][source.lower().strip()] = target.strip()
        elif mapping_type == 'visits':
            VISIT_NAME_MAPPINGS[source.lower().strip()] = target.strip()
            config['visit_mappings'][source.lower().strip()] = target.strip()
        else:
            return jsonify({'error': 'Invalid mapping type. Must be "gusto" or "visits"'}), 400
        
        save_config(config)
        return jsonify({'success': True, 'message': 'Mapping added'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/validate', methods=['POST'])
def validate_files():
    """API endpoint to validate files before submission."""
    results = {}
    
    for file_name in ['doxy_file', 'account_file', 'gusto_file', 'booking_file']:
        if file_name in request.files and request.files[file_name].filename != '':
            file_obj = request.files[file_name]
            errors = validate_file(file_obj, FILE_CONFIGS[file_name])
            
            file_obj.seek(0, 2)
            size = file_obj.tell()
            file_obj.seek(0)
            
            results[file_name] = {
                'valid': len(errors) == 0,
                'filename': file_obj.filename,
                'size': size,
                'errors': errors
            }
        else:
            results[file_name] = {
                'valid': False,
                'errors': ['No file uploaded']
            }
    
    return jsonify(results)


@app.route('/api/validate-file', methods=['POST'])
def validate_uploaded_file():
    """
    Validate an uploaded file before full processing.
    Returns validation results including errors, warnings, and file preview.
    """
    try:
        file = request.files.get('file')
        file_type = request.form.get('file_type')
        
        if not file or not file_type:
            return jsonify({'error': 'Missing file or file_type'}), 400
        
        errors = []
        warnings = []
        info = []
        
        # Read sample of file (first 50 rows for validation)
        file.seek(0)
        filename = file.filename.lower()
        
        # For Gusto, skip header rows
        skiprows = 8 if file_type == 'gusto' else 0
        
        try:
            if filename.endswith('.csv'):
                df = pd.read_csv(file, skiprows=skiprows, nrows=50)
            elif filename.endswith(('.xls', '.xlsx')):
                df = pd.read_excel(file, skiprows=skiprows, nrows=50)
            else:
                return jsonify({
                    'valid': False,
                    'errors': ['Invalid file type. Must be CSV, XLS, or XLSX'],
                    'warnings': [],
                    'info': []
                }), 400
            
            info.append(f"Successfully read file: {len(df)} rows, {len(df.columns)} columns")
            
        except Exception as e:
            logger.exception(f"Error reading file for validation: {file.filename}")
            error_msg = str(e)
            error_type = type(e).__name__
            
            # Provide more helpful error messages
            if 'No columns to parse' in error_msg or 'EmptyDataError' in error_type:
                error_msg = "File appears to be empty or has no readable data"
            elif 'skiprows' in error_msg.lower() or 'row' in error_msg.lower():
                if file_type == 'gusto':
                    error_msg = f"Could not read file structure. If this is a Gusto file, ensure it has the correct format with header rows. Error: {error_msg}"
                else:
                    error_msg = f"Could not read file structure. Error: {error_msg}"
            elif 'UnicodeDecodeError' in error_type or 'encoding' in error_msg.lower():
                error_msg = f"File encoding issue. Try saving the file as UTF-8 CSV or Excel format. Error: {error_msg}"
            elif 'BadZipFile' in error_type or 'zip' in error_msg.lower():
                error_msg = "File appears to be corrupted or not a valid Excel file. Try re-saving the file."
            elif 'xlrd' in error_msg.lower() or 'xlsx' in error_msg.lower():
                error_msg = f"Error reading Excel file. Ensure the file is not password-protected and is a valid .xls or .xlsx file. Error: {error_msg}"
            else:
                # Generic error with more context
                error_msg = f"Error reading file: {error_msg}"
            
            return jsonify({
                'valid': False,
                'errors': [error_msg],
                'warnings': [],
                'info': []
            }), 400
        
        # Type-specific validation
        if file_type == 'doxy':
            # Check required columns
            required_cols = ['Provider name', 'Duration', 'Date']
            for col in required_cols:
                # Check for exact match or close match
                if col not in df.columns:
                    # Check for case-insensitive match
                    matches = [c for c in df.columns if col.lower() in c.lower()]
                    if not matches:
                        errors.append(f"Missing required column: '{col}'")
                    else:
                        info.append(f"Found column '{matches[0]}' for '{col}'")
            
            # Validate date parsing
            if 'Date' in df.columns or any('date' in str(col).lower() for col in df.columns):
                date_col = 'Date' if 'Date' in df.columns else [c for c in df.columns if 'date' in str(c).lower()][0]
                parsed_dates = pd.to_datetime(df[date_col], errors='coerce')
                invalid_count = parsed_dates.isna().sum()
                
                if invalid_count > len(df) * 0.5:  # More than 50% invalid
                    errors.append(f"Most dates couldn't be parsed ({invalid_count}/{len(df)}). Check date format.")
                elif invalid_count > 5:
                    warnings.append(f"{invalid_count} dates couldn't be parsed. Sample: {df[parsed_dates.isna()][date_col].head(3).tolist()}")
                elif invalid_count > 0:
                    info.append(f"{invalid_count} dates couldn't be parsed")
            
            # Validate duration format
            if 'Duration' in df.columns:
                sample_durations = df['Duration'].head(10).tolist()
                valid_format = all(
                    isinstance(d, str) and d.count(':') == 2 
                    for d in sample_durations if pd.notna(d)
                )
                if not valid_format:
                    warnings.append(f"Duration format may be incorrect. Expected HH:MM:SS. Sample: {sample_durations[:3]}")
        
        elif file_type == 'gusto':
            # Validate Gusto format (should have Name, Total hours columns)
            if len(df.columns) < 4:
                errors.append(f"Unexpected format: only {len(df.columns)} columns. Expected at least 4.")
            
            # Check for name-like column
            name_cols = [c for c in df.columns if 'name' in str(c).lower()]
            if not name_cols:
                warnings.append("No column with 'name' found. First column will be used as provider names.")
            
            # Check for hours column
            hours_cols = [c for c in df.columns if 'hour' in str(c).lower() or 'total' in str(c).lower()]
            if not hours_cols:
                warnings.append("No column with 'hours' or 'total' found. May cause issues.")
        
        elif file_type == 'booking':
            # Check for booking page column (flexible matching like in actual processing)
            booking_cols = [c for c in df.columns if 'booking' in str(c).lower() or 'page' in str(c).lower() or 'provider' in str(c).lower() or 'name' in str(c).lower()]
            if not booking_cols:
                # If no booking column found, check if first column could be used
                if len(df.columns) > 0:
                    info.append("No 'Booking page' column found, but will use first column as provider name")
                else:
                    errors.append("File appears to have no columns")
            
            # Check for activity columns (these are optional - warnings only)
            expected_cols = ['All activities', 'Completed', 'Scheduled', 'Canceled', 'No-show']
            found_cols = [col for col in expected_cols if col in df.columns]
            missing = [col for col in expected_cols if col not in df.columns]
            
            if found_cols:
                info.append(f"Found activity columns: {', '.join(found_cols)}")
            if missing:
                warnings.append(f"Missing optional columns: {', '.join(missing)}. Report will still generate but may have limited data.")
        
        elif file_type == 'account':
            # Check for common Account Detail columns
            common_cols = ['provider', 'owner', 'status', 'event', 'type']
            found_cols = []
            for col in df.columns:
                col_lower = str(col).lower()
                for common in common_cols:
                    if common in col_lower:
                        found_cols.append(col)
                        break
            
            if len(found_cols) < 2:
                warnings.append(f"File format may not match expected Account Detail export. Only found columns: {', '.join(found_cols)}")
        
        # Check for completely empty rows
        if df.empty:
            errors.append("File appears to have no data rows")
        elif df.isna().all(axis=1).sum() > 0:
            empty_rows = df.isna().all(axis=1).sum()
            warnings.append(f"{empty_rows} completely empty rows found")
        
        # Reset file pointer
        file.seek(0)
        
        # Return validation results
        is_valid = len(errors) == 0
        
        return jsonify({
            'valid': is_valid,
            'errors': errors,
            'warnings': warnings,
            'info': info,
            'preview': {
                'rows': len(df),
                'columns': list(df.columns),
                'sample_data': df.head(5).fillna('').to_dict('records')
            }
        })
        
    except Exception as e:
        logger.exception("Error validating file")
        error_msg = str(e)
        # Provide more helpful error messages for common issues
        if 'No columns to parse' in error_msg or 'EmptyDataError' in str(type(e).__name__):
            error_msg = "File appears to be empty or has no readable data"
        elif 'skiprows' in error_msg.lower():
            error_msg = f"Error reading file structure. If this is a Gusto file, ensure it has the correct format with header rows. Error: {error_msg}"
        return jsonify({
            'valid': False,
            'errors': [f"Validation error: {error_msg}"],
            'warnings': [],
            'info': []
        }), 500


@app.route('/preview', methods=['POST'])
def preview_report():
    """API endpoint to preview report data before download."""
    # Validate required files
    required_files = ['doxy_file', 'account_file', 'gusto_file']
    missing_files = []
    
    for file_name in required_files:
        if file_name not in request.files or request.files[file_name].filename == '':
            missing_files.append(FILE_CONFIGS[file_name]['name'])
    
    if missing_files:
        return jsonify({'error': f"Missing required files: {', '.join(missing_files)}"}), 400
    
    try:
        doxy_file = request.files['doxy_file']
        account_file = request.files['account_file']
        gusto_file = request.files['gusto_file']
        booking_file = request.files.get('booking_file') if 'booking_file' in request.files and request.files['booking_file'].filename else None
        
        # Create diagnostic for preview (optional, won't be saved)
        diagnostic = DiagnosticLog()
        
        # Read files
        doxy_df = read_file_as_dataframe(doxy_file, diagnostic=diagnostic)
        
        account_content = None
        account_is_csv = account_file.filename.lower().endswith('.csv')
        for encoding in ['utf-16', 'utf-8', 'latin-1', 'cp1252']:
            try:
                account_file.seek(0)
                account_content = account_file.read().decode(encoding)
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        
        gusto_df = read_file_as_dataframe(gusto_file, skiprows=8, diagnostic=diagnostic)
        
        booking_df = None
        if booking_file:
            try:
                booking_df = read_file_as_dataframe(booking_file, diagnostic=diagnostic)
                booking_col = None
                for col in booking_df.columns:
                    col_lower = str(col).lower()
                    if 'booking' in col_lower or 'page' in col_lower:
                        booking_col = col
                        break
                if booking_col and booking_col != 'Booking page':
                    booking_df = booking_df.rename(columns={booking_col: 'Booking page'})
                elif 'Booking page' not in booking_df.columns:
                    booking_df = booking_df.rename(columns={booking_df.columns[0]: 'Booking page'})
            except:
                booking_df = None
        
        # Generate preview data
        doxy_visits = get_doxy_visits(doxy_df, diagnostic=diagnostic)
        doxy_providers = doxy_visits['Provider name'].tolist() if not doxy_visits.empty else []
        oncehub_visits = get_oncehub_visits(booking_df, diagnostic=diagnostic) if booking_df is not None else None
        visits_by_program = get_visits_by_program(account_content, is_csv=account_is_csv, diagnostic=diagnostic)
        gusto_hours = get_gusto_hours(gusto_df, doxy_providers, diagnostic=diagnostic)
        performance_metrics = get_doxy_performance_metrics(doxy_df, diagnostic=diagnostic)
        hours_worked = get_hours_worked(gusto_hours, visits_by_program, diagnostic=diagnostic)
        
        # Build preview response with sample data
        def get_sample_data(df, max_rows=10):
            """Get sample rows from DataFrame as list of dicts."""
            if df.empty:
                return []
            sample = df.head(max_rows)
            # Convert to dict with proper handling of NaN values
            return sample.fillna('').to_dict('records')
        
        preview = {
            'sheets': [
                {
                    'name': 'Doxy Visits',
                    'rows': len(doxy_visits),
                    'columns': list(doxy_visits.columns),
                    'sample_data': get_sample_data(doxy_visits),
                    'total_visits': int(doxy_visits['Total Visits'].sum())
                },
                {
                    'name': 'OnceHub Visits',
                    'rows': len(oncehub_visits) if oncehub_visits is not None else 0,
                    'columns': list(oncehub_visits.columns) if oncehub_visits is not None and not oncehub_visits.empty else [],
                    'sample_data': get_sample_data(oncehub_visits) if oncehub_visits is not None else [],
                    'available': oncehub_visits is not None
                },
                {
                    'name': 'Visits by Program',
                    'rows': len(visits_by_program),
                    'columns': list(visits_by_program.columns) if not visits_by_program.empty else [],
                    'sample_data': get_sample_data(visits_by_program),
                    'trt_total': int(visits_by_program['TRT'].sum()) if not visits_by_program.empty and 'TRT' in visits_by_program.columns else 0,
                    'hrt_total': int(visits_by_program['HRT'].sum()) if not visits_by_program.empty and 'HRT' in visits_by_program.columns else 0
                },
                {
                    'name': 'Gusto Hours',
                    'rows': len(gusto_hours),
                    'columns': list(gusto_hours.columns) if not gusto_hours.empty else [],
                    'sample_data': get_sample_data(gusto_hours),
                    'providers_with_hours': len(gusto_hours[gusto_hours['Total hours'] != 'N/A']) if not gusto_hours.empty else 0
                },
                {
                    'name': 'Doxy Performance',
                    'rows': len(performance_metrics),
                    'columns': list(performance_metrics.columns) if not performance_metrics.empty else [],
                    'sample_data': get_sample_data(performance_metrics),
                    'avg_duration': round(performance_metrics['Avg Duration (min)'].mean(), 1) if len(performance_metrics) > 0 else 0
                },
                {
                    'name': 'Hours Worked',
                    'rows': len(hours_worked),
                    'columns': list(hours_worked.columns) if not hours_worked.empty else [],
                    'sample_data': get_sample_data(hours_worked),
                    'total_hours': round(hours_worked['Hours Worked'].sum(), 1) if not hours_worked.empty else 0
                }
            ],
            'summary': {
                'total_providers': len(doxy_visits),
                'total_doxy_visits': int(doxy_visits['Total Visits'].sum()),
                'total_hours_worked': round(hours_worked['Hours Worked'].sum(), 1) if not hours_worked.empty else 0
            }
        }
        
        return jsonify(preview)
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/report-history', methods=['GET'])
def get_report_history():
    """Get list of previously generated reports"""
    # Note: On Vercel (serverless), file system is read-only, so report history won't persist
    # This endpoint returns empty list on serverless deployments
    is_serverless = os.environ.get('VERCEL') == '1' or os.environ.get('AWS_LAMBDA_FUNCTION_NAME') or not os.path.exists('/tmp')
    
    if is_serverless:
        # On Vercel, we can't persist files, so return empty list
        logger.info("Running in serverless environment - report history not available")
        return jsonify({'reports': []})
    
    reports_dir = 'reports'
    
    if not os.path.exists(reports_dir):
        try:
            os.makedirs(reports_dir)
        except Exception:
            # Can't create directory (likely serverless)
            return jsonify({'reports': []})
        return jsonify({'reports': []})
    
    reports = []
    try:
        for filename in os.listdir(reports_dir):
            if filename.endswith('.xlsx'):
                filepath = os.path.join(reports_dir, filename)
                stat = os.stat(filepath)
                
                reports.append({
                    'filename': filename,
                    'date': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d'),
                    'size': stat.st_size,
                    'download_url': f'/download/{filename}'
                })
        
        # Sort by date descending
        reports.sort(key=lambda x: x['date'], reverse=True)
    except Exception as e:
        logger.error(f"Error reading report history: {e}")
    
    return jsonify({'reports': reports[:10]})  # Last 10 reports


@app.route('/download/<filename>')
def download_report(filename):
    """Download a previously generated report"""
    reports_dir = 'reports'
    filepath = os.path.join(reports_dir, filename)
    
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True, download_name=filename)
    else:
        return "File not found", 404


# Explicit route for static files (needed for Vercel serverless)
@app.route('/static/<path:filename>')
def serve_static(filename):
    """Serve static files explicitly for Vercel compatibility"""
    import os
    static_path = os.path.join(os.path.dirname(__file__), 'static')
    return send_from_directory(static_path, filename)


# ============================================
# COMPREHENSIVE FEATURE ENHANCEMENTS
# ============================================

@app.route('/api/export-pdf', methods=['POST'])
def export_pdf():
    """Export report as PDF"""
    try:
        from reportlab.lib.pagesizes import letter
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib import colors
        from reportlab.lib.units import inch
        import io
        
        # Get preview data from session
        from flask import session
        preview_data = session.get('last_preview')
        if not preview_data:
            return jsonify({'error': 'No report data available. Please generate a report first.'}), 400
        
        output = io.BytesIO()
        doc = SimpleDocTemplate(output, pagesize=letter, topMargin=0.5*inch, bottomMargin=0.5*inch)
        elements = []
        styles = getSampleStyleSheet()
        
        # Custom styles
        title_style = ParagraphStyle('CustomTitle', parent=styles['Title'], fontSize=24, spaceAfter=30)
        heading_style = ParagraphStyle('CustomHeading', parent=styles['Heading1'], fontSize=16, spaceAfter=12)
        
        # Title
        generation_date = session.get('report_name', datetime.now().strftime("%m-%d-%Y"))
        elements.append(Paragraph(f"Weekly Report - {generation_date}", title_style))
        elements.append(Spacer(1, 0.2*inch))
        
        # Summary stats
        if preview_data.get('summary'):
            summary = preview_data['summary']
            summary_text = f"Total Providers: {summary.get('total_providers', 0)} | "
            summary_text += f"Total Visits: {summary.get('total_doxy_visits', 0)} | "
            summary_text += f"Total Hours: {summary.get('total_hours_worked', 0)}"
            elements.append(Paragraph(summary_text, styles['Normal']))
            elements.append(Spacer(1, 0.3*inch))
        
        # Process each sheet
        for sheet in preview_data.get('sheets', []):
            if not sheet.get('data') or not sheet['data'].get('data'):
                continue
                
            elements.append(Paragraph(sheet['name'], heading_style))
            
            # Create table
            data = sheet['data']['data']
            columns = sheet['data']['columns']
            
            if data and columns:
                # Table header
                table_data = [columns]
                
                # Table rows (limit to 50 rows per sheet for PDF)
                for row in data[:50]:
                    table_data.append([str(row.get(col, '')) for col in columns])
                
                # Create table
                table = Table(table_data)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('FONTSIZE', (0, 1), (-1, -1), 8),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
                ]))
                
                elements.append(table)
                elements.append(Spacer(1, 0.2*inch))
            
            elements.append(PageBreak())
        
        doc.build(elements)
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f'Weekly_Report_{datetime.now().strftime("%Y-%m-%d")}.pdf'
        )
    except ImportError:
        return jsonify({'error': 'PDF export requires reportlab. Install with: pip install reportlab'}), 500
    except Exception as e:
        logger.error(f"PDF export error: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/export-sheet/<sheet_name>', methods=['POST'])
def export_sheet(sheet_name):
    """Export individual sheet as CSV"""
    try:
        from flask import session
        preview_data = session.get('last_preview')
        if not preview_data:
            return jsonify({'error': 'No report data available. Please generate a report first.'}), 400
        
        # Find the sheet
        sheet = None
        for s in preview_data.get('sheets', []):
            if s['name'].lower().replace(' ', '_') == sheet_name.lower().replace(' ', '_'):
                sheet = s
                break
        
        if not sheet or not sheet.get('data'):
            return jsonify({'error': f'Sheet "{sheet_name}" not found'}), 404
        
        # Convert to CSV
        import csv
        import io
        
        output = io.StringIO()
        writer = csv.writer(output)
        
        # Write header
        columns = sheet['data'].get('columns', [])
        writer.writerow(columns)
        
        # Write data
        for row in sheet['data'].get('data', []):
            writer.writerow([str(row.get(col, '')) for col in columns])
        
        csv_output = io.BytesIO()
        csv_output.write(output.getvalue().encode('utf-8'))
        csv_output.seek(0)
        
        return send_file(
            csv_output,
            mimetype='text/csv',
            as_attachment=True,
            download_name=f'{sheet_name.replace(" ", "_")}_{datetime.now().strftime("%Y-%m-%d")}.csv'
        )
    except Exception as e:
        logger.error(f"Sheet export error: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/email-report', methods=['POST'])
def email_report():
    """Email report to specified address"""
    try:
        data = request.get_json()
        email = data.get('email')
        
        if not email:
            return jsonify({'error': 'Email address required'}), 400
        
        # Email functionality would require Flask-Mail configuration
        # For now, return a placeholder
        return jsonify({
            'message': 'Email functionality requires email server configuration',
            'status': 'not_configured'
        }), 501
    except Exception as e:
        logger.error(f"Email error: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/share-report', methods=['POST'])
def share_report():
    """Generate shareable link for report"""
    try:
        from flask import session
        import hashlib
        import json
        
        preview_data = session.get('last_preview')
        if not preview_data:
            return jsonify({'error': 'No report data available'}), 400
        
        # Generate a unique ID for this report
        report_hash = hashlib.md5(
            json.dumps(preview_data, sort_keys=True).encode()
        ).hexdigest()[:12]
        
        # In a real implementation, you'd store this in a database or cache
        # For now, we'll use session storage (limited to current user session)
        session[f'shared_report_{report_hash}'] = {
            'data': preview_data,
            'created': datetime.now().isoformat(),
            'expires': (datetime.now() + timedelta(days=7)).isoformat()
        }
        
        share_url = f"{request.host_url}api/shared/{report_hash}"
        
        return jsonify({
            'share_url': share_url,
            'expires_in': '7 days',
            'report_id': report_hash
        })
    except Exception as e:
        logger.error(f"Share error: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/shared/<report_id>', methods=['GET'])
def get_shared_report(report_id):
    """Retrieve a shared report"""
    try:
        from flask import session
        shared_data = session.get(f'shared_report_{report_id}')
        
        if not shared_data:
            return jsonify({'error': 'Report not found or expired'}), 404
        
        # Check expiration
        expires = datetime.fromisoformat(shared_data['expires'])
        if datetime.now() > expires:
            session.pop(f'shared_report_{report_id}', None)
            return jsonify({'error': 'Report has expired'}), 410
        
        return jsonify(shared_data['data'])
    except Exception as e:
        logger.error(f"Get shared report error: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/templates', methods=['GET'])
def get_templates():
    """Get saved report templates"""
    try:
        config = load_config()
        templates = config.get('templates', [])
        return jsonify({'templates': templates})
    except Exception as e:
        logger.error(f"Get templates error: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/templates', methods=['POST'])
def save_template():
    """Save a report template"""
    try:
        data = request.get_json()
        template_name = data.get('name')
        template_config = data.get('config', {})
        
        if not template_name:
            return jsonify({'error': 'Template name required'}), 400
        
        config = load_config()
        if 'templates' not in config:
            config['templates'] = []
        
        # Add or update template
        existing = next((t for t in config['templates'] if t['name'] == template_name), None)
        if existing:
            existing.update(template_config)
            existing['updated'] = datetime.now().isoformat()
        else:
            config['templates'].append({
                'name': template_name,
                'config': template_config,
                'created': datetime.now().isoformat(),
                'updated': datetime.now().isoformat()
            })
        
        save_config(config)
        return jsonify({'message': 'Template saved successfully'})
    except Exception as e:
        logger.error(f"Save template error: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/batch-process', methods=['POST'])
def batch_process():
    """Process multiple date ranges"""
    try:
        data = request.get_json()
        date_ranges = data.get('date_ranges', [])
        
        if not date_ranges:
            return jsonify({'error': 'Date ranges required'}), 400
        
        results = []
        for date_range in date_ranges:
            # This would process each date range
            # For now, return a placeholder
            results.append({
                'date_range': date_range,
                'status': 'pending',
                'message': 'Batch processing requires file uploads for each range'
            })
        
        return jsonify({'results': results, 'message': 'Batch processing initiated'})
    except Exception as e:
        logger.error(f"Batch process error: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/analytics/trends', methods=['POST'])
def get_analytics_trends():
    """Get provider performance trends"""
    try:
        from flask import session
        preview_data = session.get('last_preview')
        if not preview_data:
            return jsonify({'error': 'No report data available'}), 400
        
        # Calculate trends (simplified - would need historical data for real trends)
        trends = {
            'top_providers': [],
            'program_distribution': {},
            'visit_duration_stats': {}
        }
        
        # Get top providers
        doxy_sheet = next((s for s in preview_data['sheets'] if s['name'] == 'Doxy Visits'), None)
        if doxy_sheet and doxy_sheet.get('data', {}).get('data'):
            providers = sorted(
                doxy_sheet['data']['data'],
                key=lambda x: int(x.get('Total Visits', 0)),
                reverse=True
            )[:10]
            trends['top_providers'] = [
                {'name': p.get('Provider', ''), 'visits': int(p.get('Total Visits', 0))}
                for p in providers
            ]
        
        return jsonify(trends)
    except Exception as e:
        logger.error(f"Analytics trends error: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/error-log', methods=['GET'])
def get_error_log():
    """Export error log"""
    try:
        from flask import session
        diag = session.get('last_preview', {}).get('diagnostics', {})
        
        errors = diag.get('errors', [])
        warnings = diag.get('warnings', [])
        
        log_data = {
            'timestamp': datetime.now().isoformat(),
            'errors': errors,
            'warnings': warnings,
            'stats': diag.get('stats', {})
        }
        
        import json
        output = io.BytesIO()
        output.write(json.dumps(log_data, indent=2).encode('utf-8'))
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/json',
            as_attachment=True,
            download_name=f'error_log_{datetime.now().strftime("%Y-%m-%d_%H%M%S")}.json'
        )
    except Exception as e:
        logger.error(f"Error log export error: {e}")
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5000)
