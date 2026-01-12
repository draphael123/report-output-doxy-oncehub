"""
Weekly Report Generator - Local Flask Application
Upload spreadsheet files and generate consolidated Excel reports.
"""

from flask import Flask, render_template, request, send_file, flash, redirect, url_for, Response, jsonify
import pandas as pd
from bs4 import BeautifulSoup
import re
import io
import tempfile
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'weekly-report-generator-secret-key-2026'

# Names to exclude from all reports
EXCLUDED_NAMES = ['daniel raphael', 'dan raphael', 'draphael']

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


def get_doxy_visits(doxy_df):
    """Section 1: Count visits per provider from Doxy Report."""
    # Filter out excluded names
    doxy_df = doxy_df[~doxy_df['Provider name'].apply(should_exclude_name)]
    visits = doxy_df.groupby("Provider name").size().reset_index(name="Total Visits")
    visits = visits.sort_values("Total Visits", ascending=False)
    return visits


def get_oncehub_visits(booking_df):
    """Section 2: Get visit counts from OnceHub Booking Summary."""
    booking_df['Provider'] = booking_df['Booking page'].str.replace(r'\s*\([^)]*\)', '', regex=True).str.strip()
    # Filter out excluded names
    booking_df = booking_df[~booking_df['Provider'].apply(should_exclude_name)]
    result = booking_df[['Provider', 'All activities', 'Scheduled', 'Completed', 'Canceled', 'No-show']].copy()
    result.columns = ['Provider', 'Total Activities', 'Scheduled', 'Completed', 'Canceled', 'No-show']
    result = result.sort_values('Total Activities', ascending=False)
    return result


def get_visits_by_program(account_content, is_csv=False):
    """Section 3: Parse AccountDetailReport and categorize visits."""
    if is_csv:
        # Parse as CSV
        df = pd.read_csv(io.StringIO(account_content))
        # Map columns - adjust based on your CSV structure
        # Expected columns: Status, Owner/Provider, Event Type
        col_mapping = {}
        for col in df.columns:
            col_lower = col.lower()
            if 'status' in col_lower:
                col_mapping['Status'] = col
            elif 'owner' in col_lower or 'provider' in col_lower:
                col_mapping['Provider'] = col
            elif 'event' in col_lower and 'type' in col_lower:
                col_mapping['Event Type'] = col
            elif 'type' in col_lower:
                col_mapping['Event Type'] = col
        
        # Rename columns
        df = df.rename(columns={v: k for k, v in col_mapping.items()})
        
        # Ensure required columns exist
        if 'Status' not in df.columns:
            df['Status'] = 'Completed'
        if 'Provider' not in df.columns:
            # Try to find a name-like column
            for col in df.columns:
                if 'name' in col.lower():
                    df['Provider'] = df[col]
                    break
        if 'Event Type' not in df.columns:
            df['Event Type'] = 'Other'
    else:
        # Parse as HTML (XLS files from OnceHub are actually HTML)
        soup = BeautifulSoup(account_content, 'html.parser')
        rows = soup.find_all('tr')
        
        data = []
        for row in rows:
            cells = row.find_all('td')
            if len(cells) >= 7:
                first_cell = cells[0]
                if first_cell.get('style') and 'border-style:solid' in first_cell.get('style', ''):
                    status = cells[3].get_text(strip=True)
                    owner = cells[5].get_text(strip=True)
                    event_type = cells[6].get_text(strip=True)
                    
                    data.append({
                        'Status': status,
                        'Provider': owner,
                        'Event Type': event_type
                    })
        
        df = pd.DataFrame(data)
    
    if df.empty:
        return pd.DataFrame(columns=['Provider', 'TRT', 'HRT', 'Other', 'Total'])
    
    df = df[df['Status'] == 'Completed']
    
    # Filter out excluded names
    df = df[~df['Provider'].apply(should_exclude_name)]
    
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
    
    pivot = df.pivot_table(
        index='Provider',
        columns='Category',
        aggfunc='size',
        fill_value=0
    )
    
    pivot = pivot.reset_index()
    
    cols_order = ['Provider']
    for col in ['TRT', 'HRT', 'Other']:
        if col in pivot.columns:
            cols_order.append(col)
    pivot = pivot[cols_order]
    
    numeric_cols = [col for col in pivot.columns if col != 'Provider']
    pivot['Total'] = pivot[numeric_cols].sum(axis=1)
    pivot = pivot.sort_values('Total', ascending=False)
    
    return pivot


def get_gusto_hours(gusto_df, doxy_providers):
    """Section 4: Extract Gusto hours for providers in visit data."""
    if len(gusto_df.columns) >= 4:
        gusto_df.columns = ['Name', 'Title', 'Manager', 'Total hours'] + list(gusto_df.columns[4:])
    
    gusto_df['Name'] = gusto_df['Name'].astype(str).str.strip().str.replace('"', '')
    
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
    
    filtered = gusto_df[gusto_df['In_Doxy']][['Name', 'Total hours']].copy()
    filtered['Total hours'] = pd.to_numeric(filtered['Total hours'], errors='coerce')
    filtered = filtered[filtered['Total hours'] > 0]
    
    # Filter out excluded names
    filtered = filtered[~filtered['Name'].apply(should_exclude_name)]
    
    filtered = filtered.sort_values('Total hours', ascending=False)
    
    return filtered


def get_doxy_performance_metrics(doxy_df):
    """Section 5: Calculate performance metrics from Doxy Report."""
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
    
    metrics = metrics.sort_values('Total Visits', ascending=False)
    
    return metrics


def get_hours_worked(gusto_hours, performance_metrics):
    """Section 6: Calculate hours worked assuming all visits are 20 minutes."""
    def normalize_name(name):
        if pd.isna(name):
            return ''
        name = str(name).strip().lower()
        name = re.sub(r'\s+(np|fnp-c|md|pa|llc|inc\.?|pllc)$', '', name, flags=re.IGNORECASE)
        name = re.sub(r',\s*np$', '', name, flags=re.IGNORECASE)
        return name.strip()
    
    gusto = gusto_hours.copy()
    metrics = performance_metrics.copy()
    
    gusto['Name_norm'] = gusto['Name'].apply(normalize_name)
    metrics['Name_norm'] = metrics['Provider'].apply(normalize_name)
    
    merged = pd.merge(gusto, metrics, on='Name_norm', how='inner')
    merged['Hours Worked'] = (merged['Total Visits'] * 20 / 60).round(2)
    
    result = merged[['Name', 'Total hours', 'Total Visits', 'Hours Worked']].copy()
    result.columns = ['Provider', 'Gusto Hours', 'Total Visits', 'Hours Worked (20 min/visit)']
    result = result.sort_values('Gusto Hours', ascending=False)
    
    return result


def read_file_as_dataframe(file_obj, skiprows=0):
    """Read a file as DataFrame, handling both CSV and XLS formats."""
    filename = file_obj.filename.lower()
    is_excel = filename.endswith('.xls') or filename.endswith('.xlsx')
    
    file_obj.seek(0)
    
    if is_excel:
        # Try reading as Excel first
        try:
            return pd.read_excel(file_obj, skiprows=skiprows)
        except Exception:
            # If that fails, it might be HTML disguised as XLS (OnceHub style)
            file_obj.seek(0)
            content = None
            for encoding in ['utf-16', 'utf-8', 'latin-1', 'cp1252']:
                try:
                    file_obj.seek(0)
                    content = file_obj.read().decode(encoding)
                    break
                except (UnicodeDecodeError, UnicodeError):
                    continue
            if content:
                # Try to read as HTML table
                try:
                    tables = pd.read_html(io.StringIO(content))
                    if tables:
                        df = tables[0]
                        if skiprows > 0:
                            df = df.iloc[skiprows:]
                            df.columns = df.iloc[0]
                            df = df.iloc[1:].reset_index(drop=True)
                        return df
                except Exception:
                    pass
            raise ValueError("Could not read Excel file")
    else:
        # Read as CSV
        return pd.read_csv(file_obj, skiprows=skiprows)


def generate_report(doxy_file, account_file, gusto_file, booking_file):
    """Generate the complete Excel report with detailed error handling."""
    errors = []
    
    # Read Doxy Report (CSV or XLS)
    try:
        doxy_df = read_file_as_dataframe(doxy_file)
        if 'Provider name' not in doxy_df.columns:
            errors.append("Doxy Report missing 'Provider name' column")
        if 'Duration' not in doxy_df.columns:
            errors.append("Doxy Report missing 'Duration' column")
    except Exception as e:
        errors.append(f"Error reading Doxy Report: {str(e)}")
        doxy_df = None
    
    # Read Account Detail Report (try different encodings)
    account_content = None
    account_is_csv = account_file.filename.lower().endswith('.csv')
    
    for encoding in ['utf-16', 'utf-8', 'latin-1', 'cp1252']:
        try:
            account_file.seek(0)
            account_content = account_file.read().decode(encoding)
            break
        except (UnicodeDecodeError, UnicodeError):
            continue
    
    if account_content is None:
        errors.append("Could not decode Account Detail Report - try a different file format")
    
    # Read Gusto Hours (CSV or XLS, skip header rows)
    try:
        gusto_df = read_file_as_dataframe(gusto_file, skiprows=8)
    except Exception as e:
        errors.append(f"Error reading Gusto file: {str(e)}")
        gusto_df = None
    
    # Read Booking Summary (OnceHub) - CSV or XLS (OPTIONAL)
    booking_df = None
    if booking_file and booking_file.filename:
        try:
            booking_df = read_file_as_dataframe(booking_file)
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
    doxy_visits = get_doxy_visits(doxy_df)
    doxy_providers = doxy_visits['Provider name'].tolist()
    
    # OnceHub is optional
    oncehub_visits = get_oncehub_visits(booking_df) if booking_df is not None else None
    visits_by_program = get_visits_by_program(account_content, is_csv=account_is_csv)
    gusto_hours = get_gusto_hours(gusto_df, doxy_providers)
    performance_metrics = get_doxy_performance_metrics(doxy_df)
    hours_worked = get_hours_worked(gusto_hours, performance_metrics)
    
    # Calculate stats for response
    stats = {
        'providers': len(doxy_visits),
        'total_visits': int(doxy_visits['Total Visits'].sum()),
        'sheets': 6
    }
    
    # Create Excel file in memory
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        doxy_visits.to_excel(writer, sheet_name='Doxy Visits', index=False)
        if oncehub_visits is not None:
            oncehub_visits.to_excel(writer, sheet_name='OnceHub Visits', index=False)
        visits_by_program.to_excel(writer, sheet_name='Visits by Program', index=False)
        gusto_hours.to_excel(writer, sheet_name='Gusto Hours', index=False)
        performance_metrics.to_excel(writer, sheet_name='Doxy Performance Metrics', index=False)
        hours_worked.to_excel(writer, sheet_name='Hours Worked', index=False)
    
    output.seek(0)
    return output, stats


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
            output, stats = generate_report(doxy_file, account_file, gusto_file, booking_file)
            
            # Get report name from form or generate from dates
            report_name = request.form.get('report_name', '').strip()
            
            if not report_name:
                start_date = request.form.get('start_date', '')
                end_date = request.form.get('end_date', '')
                
                if start_date and end_date:
                    try:
                        start = datetime.strptime(start_date, '%Y-%m-%d')
                        end = datetime.strptime(end_date, '%Y-%m-%d')
                        report_name = f"Report {start.month}-{start.day} to {end.month}-{end.day}"
                    except ValueError:
                        report_name = 'Weekly Report'
                else:
                    report_name = f'Weekly Report {datetime.now().strftime("%m-%d-%Y")}'
            
            return Response(
                output.getvalue(),
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                headers={
                    'Content-Disposition': f'attachment; filename="{report_name}.xlsx"',
                    'X-Report-Providers': str(stats['providers']),
                    'X-Report-Visits': str(stats['total_visits'])
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


if __name__ == '__main__':
    app.run(debug=True, port=5000)
