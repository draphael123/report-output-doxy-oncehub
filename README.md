# Weekly Report Generator

A web application to generate consolidated Excel reports from healthcare provider data files.

## Features

Upload 4 data files and automatically generate a comprehensive Excel report with 6 sheets:

1. **Doxy Visits** - Visit counts per provider from Doxy
2. **OnceHub Visits** - Visit counts from OnceHub booking system
3. **Visits by Program** - Grouped by TRT, HRT, and Other
4. **Gusto Hours** - Hours worked from time tracking
5. **Doxy Performance Metrics** - Visits over 20 minutes analysis
6. **Hours Worked** - Calculated hours assuming 20 min/visit

## Required Input Files

| File | Format | Description |
|------|--------|-------------|
| Doxy Report | CSV | Provider visit data with durations |
| Account Detail Report | XLS | Visit details with program types |
| Gusto Hours | CSV | Time tracking hours per provider |
| OnceHub Booking Summary | CSV | Booking page activity summary |

## Local Development

### Prerequisites
- Python 3.9+
- pip

### Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/weekly-report-generator.git
cd weekly-report-generator

# Install dependencies
pip install -r requirements.txt

# Run locally
python app.py
```

Then open http://localhost:5000 in your browser.

## Deploy to Vercel

### One-Click Deploy

[![Deploy with Vercel](https://vercel.com/button)](https://vercel.com/new/clone?repository-url=https://github.com/yourusername/weekly-report-generator)

### Manual Deploy

1. Install Vercel CLI:
```bash
npm i -g vercel
```

2. Deploy:
```bash
vercel
```

## Project Structure

```
webapp/
├── api/
│   └── index.py          # Serverless function for Vercel
├── templates/
│   └── index.html        # Web interface
├── app.py                # Local development server
├── requirements.txt      # Python dependencies
├── vercel.json          # Vercel configuration
└── README.md            # This file
```

## Report Specifications

### Doxy Visits
- Source: `Doxy Report.csv`
- Groups visits by provider name
- Sorted by total visits (descending)

### OnceHub Visits
- Source: `BookingPageSummaryReport.csv`
- Shows: Total Activities, Scheduled, Completed, Canceled, No-show
- Sorted by total activities (descending)

### Visits by Program
- Source: `AccountDetailReport.xls`
- Categorizes by: TRT, HRT, Other
- Filters to Completed status only
- Combines Initial and Follow-up visits

### Gusto Hours
- Source: Time tracking CSV
- Only includes providers who appear in Doxy data
- Filters out providers with 0 hours

### Doxy Performance Metrics
- Calculates visits over 20 minutes
- Shows percentage and average duration
- Includes hours spent on 20+ min visits

### Hours Worked
- Combines Gusto hours with visit data
- Calculates expected hours: `Total Visits × 20 min ÷ 60`
- Allows comparison of tracked vs calculated hours

## License

MIT License - feel free to use and modify.

