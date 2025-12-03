# 📊 Schedule Validator

A web-based tool to compare CSV schedule data with PDF files and identify discrepancies.

## 🎯 Features

- **CSV Upload**: Upload schedule data in CSV format (same structure as Google Sheets)
- **PDF Upload**: Upload Kemps Corner and/or Bandra PDF schedules
- **Smart Normalization**: Automatically normalizes class names, trainer names, and times
- **Cover Replacement Support**: Handles trainer cover replacements from CSV
- **Detailed Comparison**: Shows side-by-side comparison of CSV vs PDF data
- **Summary Statistics**: Displays total classes, matches, discrepancies, and accuracy rate
- **Beautiful UI**: Modern, responsive interface with drag-and-drop support

## 📋 Requirements

- Python 3.7+
- PyPDF2 library

## 🚀 Installation

1. Install the required Python package:
```bash
pip install -r validator_requirements.txt
```

## 💻 Usage

1. Start the validation server:
```bash
python schedule_validator.py
```

2. Open your browser and navigate to:
```
http://localhost:8080
```

3. Upload your files:
   - **CSV File**: Required - Upload your schedule CSV with columns: Location, Day, Time, Class, Trainer, (optional) Cover Trainer
   - **Kemps PDF**: Optional - Upload Kemps Corner schedule PDF
   - **Bandra PDF**: Optional - Upload Bandra schedule PDF

4. Click "Validate Schedules" to start the comparison

5. View the results:
   - Summary statistics showing matches and discrepancies
   - Detailed list of all discrepancies found
   - Side-by-side comparison of CSV vs PDF data

## 📝 CSV Format

Your CSV file should have the following columns:

| Location | Day | Time | Class | Trainer | Cover Trainer (optional) |
|----------|-----|------|-------|---------|--------------------------|
| KEMPS | MONDAY | 7:15 AM | STRENGTH LAB (PULL) | ANISHA | |
| BANDRA | TUESDAY | 9:00 AM | BARRE 57 | KARAN | SIMRAN |

**Column Descriptions:**
- `Location`: KEMPS or BANDRA
- `Day`: MONDAY, TUESDAY, WEDNESDAY, THURSDAY, FRIDAY, SATURDAY, SUNDAY
- `Time`: Class time in format "HH:MM AM/PM"
- `Class`: Class name
- `Trainer`: Regular trainer name
- `Cover Trainer`: (Optional) Replacement trainer if covering

## 🔧 Normalization Features

The validator automatically normalizes:

### Class Names
- Removes extra spaces and standardizes formatting
- Handles common aliases (e.g., "CARDIO BARRE" = "CARDIOBARRE" = "CB")
- Removes parentheses and special characters

### Trainer Names
- Converts to uppercase
- Removes extra spaces
- Standardizes formatting

### Times
- Supports multiple formats: "9:00 AM", "9.00 AM", "9:00AM", "9 AM"
- Converts all to standard "H:MM AM/PM" format

### Cover Replacements
- If "Cover Trainer" column is provided in CSV, uses that trainer instead of regular trainer
- Useful for temporary trainer replacements

## 📊 Results Interpretation

### Summary Statistics
- **Total Classes**: All classes found in CSV and PDF
- **Matched**: Classes that match exactly between CSV and PDF
- **Discrepancies**: Classes with differences
- **Accuracy**: Percentage of matches (Matched / Total * 100%)

### Discrepancy Types
1. **Class Mismatch**: Different class names
2. **Trainer Mismatch**: Different trainer names
3. **Not in CSV**: Class exists in PDF but not in CSV
4. **Not in PDF**: Class exists in CSV but not in PDF

## 🎨 Interface Features

- **Drag and Drop**: Drag files directly onto upload boxes
- **Click to Browse**: Traditional file browsing
- **Visual Feedback**: Clear indication of uploaded files
- **Responsive Design**: Works on desktop and mobile
- **Color-Coded Results**: 
  - 🟢 Green: Matches
  - 🔴 Red: Discrepancies
  - ⚪ Gray: Neutral info

## 🛠️ Troubleshooting

### Server won't start
- Make sure port 8080 is not already in use
- Check that Python 3.7+ is installed
- Verify PyPDF2 is installed

### PDF parsing issues
- Ensure PDFs are text-based (not scanned images)
- Check that PDF format matches expected layout
- PDFs should have clear day headers and time/class/trainer format

### CSV parsing issues
- Verify CSV has correct column headers
- Check for missing required columns
- Ensure proper CSV formatting (commas, quotes)

## 📞 Support

For issues or questions, check the application logs in the terminal where the server is running.

## 🔒 Security Note

This tool runs locally on your machine. No data is sent to external servers.

## 📄 License

For internal use only.
