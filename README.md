## Support linux chrome browser

# BRACU Connect Auto Mark Entry - Extension

A Chrome extension that automatically fills student marks in BRACU Connect Final Mark Entry system by importing data from Excel files.

## Features

- **Smart Page Detection**: Only works on BRACU Connect Final Mark Entry pages
- **Excel Import**: Upload .xlsx files with student marks
- **Automatic Matching**: Matches students by ID and fills marks instantly  
- **Error Reporting**: Shows unmatched students and validation errors
- **Visual Feedback**: Highlights successfully filled fields

## User Guidelines

### Excel File Requirements
- File must be `.xlsx` format
- Sheet name must be "Final GradeSheet"
- Must contain "ID #" column and "Total" column
- Student data should start 2 rows after header

## Installation

### Method 1: Manual Installation
1. Download this repository
2. Open Chrome → `chrome://extensions/`
3. Enable "Developer mode"
4. Click "Load unpacked" → Select extension folder

### Method 2: Chrome Web Store
*(Coming soon)*

## How to Use

1. **Navigate to BRACU Connect Final Mark Entry page**
2. **Click the extension icon** in Chrome toolbar
3. **Upload Excel file** (.xlsx with "Final GradeSheet" sheet)
4. **Click "Fill Marks"** button
5. **Review results** - green highlights show successful entries

### Expected Results:
- ✅ All matched students get marks filled automatically
- ⚠️ Unmatched students are listed for manual entry
- ❌ Error messages guide you if something goes wrong

## Technical Details

**Built with**: JavaScript, Chrome Extension APIs, SheetJS
**File Processing**: Reads .xlsx files, finds ID# and Total columns dynamically
**DOM Interaction**: Locates StudentId and Marks input fields, fills and triggers events
**Error Handling**: Validates page, file format, and data matching

## Security & Privacy

This extension is **completely safe** and poses **no security threats**:

- **No Data Collection**: Extension doesn't collect, store, or transmit any personal data
- **Local Processing**: All Excel processing happens locally in your browser
- **No External Servers**: No data is sent to external servers or third parties
- **Limited Permissions**: Only accesses active tab when you explicitly use the extension
- **Open Source**: All code is visible and auditable
- **No Network Requests**: Extension works entirely offline after loading

**Privacy Guarantee**: Your student data, marks, and Excel files never leave your computer.

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "Wrong Page Detected" | Navigate to BRACU Connect Final Mark Entry page |
| "No Students Matched" | Check if course/section matches between Excel and page |
| File won't upload | Ensure .xlsx format with "Final GradeSheet" sheet |
| Extension not working | Refresh page and try again |

## Contributing

1. Fork repository
2. Make changes
3. Test with real BRACU Connect pages
4. Submit pull request

## License

MIT License - Free to use and modify

---

**Made for BRACU Community** 🎓
