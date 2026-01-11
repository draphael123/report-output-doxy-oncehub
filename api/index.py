"""
Weekly Report Generator - Vercel Serverless Function
Upload spreadsheet files and generate consolidated Excel reports.
"""

from flask import Flask, render_template, request, send_file, flash, redirect, url_for, Response
import pandas as pd
from bs4 import BeautifulSoup
import re
import io
import tempfile
from datetime import datetime

app = Flask(__name__, template_folder='../templates')
app.secret_key = 'weekly-report-generator-secret-key-2026'


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
    visits = doxy_df.groupby("Provider name").size().reset_index(name="Total Visits")
    visits = visits.sort_values("Total Visits", ascending=False)
    return visits


def get_oncehub_visits(booking_df):
    """Section 2: Get visit counts from OnceHub Booking Summary."""
    booking_df['Provider'] = booking_df['Booking page'].str.replace(r'\s*\([^)]*\)', '', regex=True).str.strip()
    result = booking_df[['Provider', 'All activities', 'Scheduled', 'Completed', 'Canceled', 'No-show']].copy()
    result.columns = ['Provider', 'Total Activities', 'Scheduled', 'Completed', 'Canceled', 'No-show']
    result = result.sort_values('Total Activities', ascending=False)
    return result


def get_visits_by_program(account_content):
    """Section 3: Parse AccountDetailReport and categorize visits."""
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
    filtered = filtered.sort_values('Total hours', ascending=False)
    
    return filtered


def get_doxy_performance_metrics(doxy_df):
    """Section 5: Calculate performance metrics from Doxy Report."""
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


def generate_report(doxy_file, account_file, gusto_file, booking_file):
    """Generate the complete Excel report."""
    # Read Doxy Report
    doxy_df = pd.read_csv(doxy_file)
    
    # Read Account Detail Report (try different encodings)
    account_content = None
    for encoding in ['utf-16', 'utf-8', 'latin-1', 'cp1252']:
        try:
            account_file.seek(0)
            account_content = account_file.read().decode(encoding)
            break
        except (UnicodeDecodeError, UnicodeError):
            continue
    
    if account_content is None:
        raise ValueError("Could not decode Account Detail Report")
    
    # Read Gusto Hours (skip header rows)
    gusto_file.seek(0)
    gusto_df = pd.read_csv(gusto_file, skiprows=8, header=0)
    
    # Read Booking Summary (OnceHub)
    booking_file.seek(0)
    booking_df = pd.read_csv(booking_file)
    
    # Generate all sections
    doxy_visits = get_doxy_visits(doxy_df)
    doxy_providers = doxy_visits['Provider name'].tolist()
    
    oncehub_visits = get_oncehub_visits(booking_df)
    visits_by_program = get_visits_by_program(account_content)
    gusto_hours = get_gusto_hours(gusto_df, doxy_providers)
    performance_metrics = get_doxy_performance_metrics(doxy_df)
    hours_worked = get_hours_worked(gusto_hours, performance_metrics)
    
    # Create Excel file in memory
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        doxy_visits.to_excel(writer, sheet_name='Doxy Visits', index=False)
        oncehub_visits.to_excel(writer, sheet_name='OnceHub Visits', index=False)
        visits_by_program.to_excel(writer, sheet_name='Visits by Program', index=False)
        gusto_hours.to_excel(writer, sheet_name='Gusto Hours', index=False)
        performance_metrics.to_excel(writer, sheet_name='Doxy Performance Metrics', index=False)
        hours_worked.to_excel(writer, sheet_name='Hours Worked', index=False)
    
    output.seek(0)
    return output


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if all files are present
        required_files = ['doxy_file', 'account_file', 'gusto_file', 'booking_file']
        for file_name in required_files:
            if file_name not in request.files or request.files[file_name].filename == '':
                flash('Please upload all required files', 'error')
                return redirect(request.url)
        
        doxy_file = request.files['doxy_file']
        account_file = request.files['account_file']
        gusto_file = request.files['gusto_file']
        booking_file = request.files['booking_file']
        
        try:
            # Generate the report
            output = generate_report(doxy_file, account_file, gusto_file, booking_file)
            
            # Get report name from form or use default
            report_name = request.form.get('report_name', 'Weekly Report')
            if not report_name:
                report_name = 'Weekly Report'
            
            return Response(
                output.getvalue(),
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                headers={
                    'Content-Disposition': f'attachment; filename="{report_name}.xlsx"'
                }
            )
        except Exception as e:
            flash(f'Error generating report: {str(e)}', 'error')
            return redirect(request.url)
    
    return render_template('index.html')


# Vercel serverless handler
app = app

