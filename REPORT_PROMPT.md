# Weekly Report Generation Prompt

## Task
Create a document named "Report 1-11" in Excel format (.xlsx). This document should act as a consolidated summary report based on data from the following spreadsheets:

- `Doxy Report.csv` - Provider visit data with durations
- `AccountDetailReport_[date range].xls` - Visit details with program types
- `duval-medical-p-a-time-tracking-hours-[date range].csv` - Gusto hours data
- `BookingPageSummaryReport_[date range].csv` - OnceHub booking summary

---

## Report Structure (6 Sheets)

### 1. Doxy Visits
- List the number of visits each provider had during the date range
- Source: Doxy Report.csv
- Columns: Provider, Total Visits

### 2. OnceHub Visits
- List the number of visits from the OnceHub booking system
- Source: BookingPageSummaryReport.csv
- Columns: Provider, Total Activities, Scheduled, Completed, Canceled, No-show

### 3. Visits by Program
- Group visits by category: TRT, HRT, or Other
- Combine all visit types (Initial and Follow-up) together - do NOT separate by visit type
- Source: AccountDetailReport (filter to Completed status only)
- Columns: Provider, TRT, HRT, Other, Total

### 4. Gusto Hours
- Extract total hours worked for each provider from the Gusto/time-tracking file
- Only include providers who also appear in the Doxy/visit-related data
- Columns: Name, Total Hours

### 5. Doxy Performance Metrics (Visits over 20 mins)
For each provider, include:
- Total number of visits
- Number of visits exceeding 20 minutes
- Percentage of total visits that lasted over 20 minutes
- Hours spent on visits over 20 minutes
- Average visit duration in minutes
- Source: Doxy Report.csv (parse Duration column)

### 6. Hours Worked
- Combine Gusto hours with calculated visit hours
- Calculate how long each provider actually worked assuming all visits are 20 minutes
- Formula: Total Visits × 20 minutes ÷ 60 = Hours Worked
- Columns: Provider, Gusto Hours, Total Visits, Hours Worked (20 min/visit)

---

## Data Processing Notes
- Match provider names between files using normalized name matching (handle suffixes like NP, FNP-C, MD, etc.)
- Handle different file encodings (UTF-16, UTF-8, etc.)
- Parse duration strings in HH:MM:SS format to calculate minutes
- Sort tables by relevant metrics (typically descending by total visits or hours)
- Present data in clean, tabular format for easy reading

---

## Required Files
1. **Doxy Report** (CSV) - Contains: Date, Provider name, Duration, etc.
2. **Account Detail Report** (XLS/HTML) - Contains: Meeting date, Subject/Event type, Status, Booking page owner
3. **Gusto Hours** (CSV) - Contains: Name, Total hours
4. **OnceHub Booking Summary** (CSV) - Contains: Booking page, All activities, Scheduled, Completed, Canceled, No-show

---

## Web Application Usage

### Local Development
```bash
cd webapp
pip install -r requirements.txt
python app.py
```
Then open http://localhost:5000

### Vercel Deployment
```bash
cd webapp
vercel
```

### GitHub Repository Setup
1. Create a new repository on GitHub
2. Push the webapp folder contents:
```bash
cd webapp
git init
git add .
git commit -m "Initial commit - Weekly Report Generator"
git remote add origin https://github.com/yourusername/weekly-report-generator.git
git push -u origin main
```

3. Connect to Vercel:
   - Go to vercel.com
   - Import your GitHub repository
   - Deploy automatically



