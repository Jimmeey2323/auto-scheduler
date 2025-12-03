# 🚀 Schedule Validator - Quick Start Guide

## What is this?

The Schedule Validator is a web-based tool that compares your CSV schedule data with PDF schedule files to find any discrepancies. It's perfect for:

- ✅ Verifying schedule accuracy before publishing
- 🔍 Finding mismatches between planned and printed schedules
- 📊 Quality control for schedule updates
- 🎯 Identifying trainer cover replacements that weren't applied

---

## 📋 Step-by-Step Usage

### Step 1: Start the Server

**Option A - Using the startup script (Recommended):**
```bash
./start_validator.sh
```

**Option B - Manual start:**
```bash
python3 schedule_validator.py
```

You should see:
```
🚀 Schedule Validator Server Starting...
📍 Server running at: http://localhost:8080
🌐 Open your browser and navigate to the URL above
```

### Step 2: Open the Interface

1. Open your web browser
2. Go to: `http://localhost:8080`
3. You'll see the Schedule Validator interface

### Step 3: Prepare Your Files

**CSV File (Required):**
- Must include columns: Location, Day, Time, Class, Trainer
- Optional column: Cover Trainer (for trainer replacements)
- Same format as your Google Sheets schedule

**Example CSV:**
```csv
Location,Day,Time,Class,Trainer,Cover Trainer
KEMPS,MONDAY,7:15 AM,STRENGTH LAB (PULL),ANISHA,
BANDRA,TUESDAY,9:00 AM,BARRE 57,SIMRAN,KARAN
```

**PDF Files (At least one required):**
- Kemps Corner PDF schedule
- Bandra PDF schedule
- Or both!

### Step 4: Upload Files

1. **CSV Upload Section:**
   - Click the CSV upload box OR drag & drop your CSV file
   - You'll see the filename once uploaded

2. **PDF Upload Section:**
   - Upload Kemps PDF (optional)
   - Upload Bandra PDF (optional)
   - At least one PDF is required

3. **Validate Button:**
   - Button enables when you have CSV + at least 1 PDF
   - Click "Validate Schedules" to start comparison

### Step 5: Review Results

The results page shows:

**📊 Summary Card:**
- Total Classes: Number of classes checked
- Matched: Classes that match perfectly
- Discrepancies: Classes with differences
- Accuracy: Percentage match rate

**📋 Discrepancy List:**
- Shows each mismatch found
- Side-by-side comparison of CSV vs PDF data
- Color-coded:
  - 🔴 Red = Mismatch
  - 🟢 Green = Match
  - ⚪ Gray = Not found

---

## 🎯 Common Use Cases

### Case 1: Weekly Schedule Check
**When:** Before publishing weekly schedules  
**How:** Upload your planned schedule CSV and current week's PDFs  
**Goal:** Ensure everything matches before client distribution

### Case 2: Trainer Cover Verification
**When:** After updating cover replacements  
**How:** Add cover trainers to CSV's "Cover Trainer" column  
**Goal:** Verify replacements are correctly reflected in PDFs

### Case 3: Location-Specific Audit
**When:** Checking one specific location  
**How:** Upload CSV with all locations, but only one PDF  
**Goal:** Focus audit on specific location

### Case 4: Post-Update Validation
**When:** After updating HTML/PDF files  
**How:** Upload fresh CSV export and newly generated PDFs  
**Goal:** Confirm all updates applied correctly

---

## 💡 Pro Tips

### 1. CSV Preparation
- Export directly from Google Sheets as CSV
- Keep the same column order
- Include Cover Trainer column even if empty
- Save with UTF-8 encoding

### 2. Cover Replacements
If a trainer is covering for another:
```csv
Location,Day,Time,Class,Trainer,Cover Trainer
KEMPS,MONDAY,9:00 AM,BARRE 57,ANISHA,ROHAN
```
The validator will check for "ROHAN" in the PDF, not "ANISHA"

### 3. Handling Discrepancies

**Class Name Mismatch:**
- Check for typos in CSV or PDF
- Verify class name aliases are recognized
- Common issue: Extra spaces or punctuation

**Trainer Name Mismatch:**
- Check spelling consistency
- Verify cover trainer is in right column
- Common issue: Nickname vs. full name

**Time Mismatch:**
- Usually indicates a scheduling change
- Verify correct time in source data
- Check if time was updated in both places

**Not Found:**
- "NOT IN CSV" = Class exists in PDF but not in schedule
- "NOT IN PDF" = Class exists in schedule but not in PDF
- Review if class was added/removed recently

### 4. Best Practices
- ✅ Run validation BEFORE distributing schedules
- ✅ Keep a log of accuracy rates over time
- ✅ Address discrepancies immediately
- ✅ Re-run validation after fixing issues
- ✅ Archive validation results for records

---

## 🔧 Troubleshooting

### Problem: Server won't start
**Solution:** Check if port 8080 is in use
```bash
lsof -i :8080
```
If in use, kill the process or change port in `schedule_validator.py`

### Problem: CSV parse errors
**Solutions:**
- Verify CSV has correct headers
- Check for missing commas
- Ensure UTF-8 encoding
- Remove empty rows at end

### Problem: PDF not parsing correctly
**Solutions:**
- Ensure PDF is text-based (not scanned image)
- Check PDF layout matches expected format
- Verify day headers are present (MONDAY, TUESDAY, etc.)
- Confirm time format is readable

### Problem: Too many false positives
**Solutions:**
- Check normalization rules in `schedule_validator.py`
- Add more aliases to `CLASS_ALIASES` dictionary
- Verify time formats are consistent

### Problem: Browser won't connect
**Solutions:**
- Verify server is running (check terminal)
- Try `http://127.0.0.1:8080` instead
- Check firewall settings
- Clear browser cache

---

## 📊 Understanding Results

### 100% Accuracy
🎉 Perfect! CSV and PDF match completely. Ready to publish.

### 90-99% Accuracy
✅ Very good. Review minor discrepancies. Likely typos or recent changes.

### 70-89% Accuracy
⚠️ Moderate issues. Several mismatches need attention. Check recent updates.

### Below 70% Accuracy
🚨 Significant problems. Major sync issues between CSV and PDF. Review thoroughly.

---

## 🔒 Privacy & Security

- ✅ Runs 100% locally on your machine
- ✅ No data sent to external servers
- ✅ Files processed in memory only
- ✅ No data stored after validation
- ✅ Safe for confidential schedule data

---

## 📞 Need Help?

1. Check the terminal where server is running for error messages
2. Run the test suite: `python3 test_validator.py`
3. Review `VALIDATOR_README.md` for detailed documentation
4. Check sample files: `sample_schedule.csv`

---

## 🎓 Example Workflow

```
1. Export schedule from Google Sheets → schedule.csv
2. Download Kemps.pdf from Google Drive
3. Download Bandra.pdf from Google Drive
4. Start validator: ./start_validator.sh
5. Open browser: http://localhost:8080
6. Upload: schedule.csv, Kemps.pdf, Bandra.pdf
7. Click: Validate Schedules
8. Review: Results and discrepancies
9. Fix: Any issues found
10. Re-run: Validation until 100%
```

---

## 🎯 Success Checklist

Before publishing schedules, ensure:

- [ ] Validation run successfully
- [ ] Accuracy rate above 95%
- [ ] All discrepancies reviewed
- [ ] Critical mismatches fixed
- [ ] Re-validation confirms fixes
- [ ] PDFs match current week
- [ ] Cover trainers all applied
- [ ] Times are correct
- [ ] Class names consistent
- [ ] Ready for distribution

---

**Happy Validating! 🚀**
