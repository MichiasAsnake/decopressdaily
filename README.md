# DecoPress Daily Tool

A tool for DecoPress to generate daily order lists and packing slips.

## Features

1. **Daily Order List**
   - Automatically scrapes urgent orders (0, 1, and 2 days remaining)
   - Creates an Excel file with a list of orders
   - Saves to Desktop/Decopress_Downloads with date in filename

2. **Packing Slip Generator**
   - Enter a job number
   - Automatically finds the job information and shipping details
   - Creates a packing slip Excel file
   - Saves to Desktop/Decopress_Downloads with date and job number in filename

## Requirements

- Python 3.7+
- Required Python packages:
  - playwright
  - pandas
  - pillow (PIL)
  - openpyxl

## Installation

1. Install required packages:
```
pip install playwright pandas openpyxl pillow
```

2. Install Playwright browsers:
```
playwright install
```

## Usage

1. Run the application:
```
python app.py
```

2. On the welcome screen, choose either:
   - "Create Daily Order List" - Scrapes urgent orders
   - "Create Packing Slip" - Generates a packing slip for a specific job

3. Login with your DecoPress credentials when prompted

4. For packing slips, enter the job number when prompted

5. The generated files will be saved to Desktop/Decopress_Downloads

## File Structure

- `app.py` - Main application with UI
- `daily_orders.py` - Daily order list scraping logic
- `packing_slip.py` - Packing slip generation logic
- `utils.py` - Shared utility functions
- `DecoPressLogo.jpg` - DecoPress logo for the UI

## Notes

- Make sure the DecoPressLogo.jpg file is in the same directory as the app.py file
- The application will create a Decopress_Downloads folder on your Desktop if it doesn't exist
- Each run will create a new file with the current date in the filename 