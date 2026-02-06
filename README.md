# Auto Scheduler

An automated system for managing and updating class schedules, integrating with Google Sheets and generating PDF schedules.

## Features

- 📅 Automated schedule updates from Google Sheets
- 🔄 HTML template processing for different locations (Bandra, Kemps)
- 📧 Email notifications with PDF attachments
- 🎨 Theme-based styling (PowerCycle, Total, Core Yoga)
- 📊 Google Sheets integration for real-time data
- 🚫 **Sold Out marking** - Set theme to "Sold Out" to add red strikethrough and badge

### Sold Out Feature

Mark classes as sold out by setting the **Theme** column to `Sold Out` in the Google Sheets "Cleaned" tab. This will:
- Add a red strikethrough line across the class row
- Display a red "SOLD OUT" badge next to the class name

See [SOLD_OUT_FEATURE.md](SOLD_OUT_FEATURE.md) for detailed usage instructions.

## Setup

### 1. Install Dependencies

```bash
npm install
```

### 2. Environment Configuration

Create a `.env` file in the project root with your Google OAuth credentials:

```bash
cp .env.example .env
```

Fill in the following environment variables in your `.env` file:

```bash
GOOGLE_CLIENT_ID=your_google_client_id
GOOGLE_CLIENT_SECRET=your_google_client_secret
GOOGLE_REFRESH_TOKEN=your_google_refresh_token
GOOGLE_FOLDER_ID=your_google_drive_folder_id
GOOGLE_SPREADSHEET_ID=your_google_spreadsheet_id
GOOGLE_SHEET_NAME=your_sheet_name
```

### 3. Google OAuth Setup

Run the Gmail authentication setup:

```bash
node setupGmailAuth.js
```

Follow the prompts to authorize the application and generate refresh tokens.

## Usage

### Full Workflow (Email + Sheets + HTML)

Processes email, updates Google Sheets, then generates HTML and PDF:

```bash
npm run update
# or
node updateKempsSchedule.js
```

### HTML-Only Update (Skip Email Processing)

Use existing data in the Cleaned sheet to update HTML and PDF without processing emails:

```bash
npm run update:html-only
# or
npm run update:skip-email
# or
node updateKempsSchedule.js --skip-email
```

**When to use HTML-only mode:**
- Data is already in the Cleaned sheet
- You want to regenerate HTML/PDF without re-processing emails
- Making manual adjustments to the Cleaned sheet
- Faster updates when email data hasn't changed

### Help and Options

View all available options:

```bash
node updateKempsSchedule.js --help
```

### Legacy Scripts

Update HTML only (deprecated - use `--skip-email` flag instead):
```bash
node updateHTMLOnly.js
```

Complete schedule update:
```bash
node completeScheduleUpdate.js
```

## File Structure

- `Bandra.html` / `Kemps.html` - Schedule template files
- `updateHTMLOnly.js` - Updates HTML templates with latest data
- `updateKempsSchedule.js` - Full schedule processing with email
- `setupGmailAuth.js` - Google OAuth configuration helper
- `package.json` - Project dependencies and scripts

## Features

### Supported Themes
- PowerCycle Monday (Blue theme)
- Total (Red theme) 
- Core Yoga (Green theme)
- Default theme for other classes

### Automated Processing
- Fetches data from Google Sheets
- Processes schedule information
- Updates HTML templates
- Generates PDF outputs
- Sends email notifications

## Security

Sensitive credentials are stored in environment variables and not committed to the repository. Make sure to:

1. Never commit your `.env` file
2. Use secure OAuth tokens
3. Regularly rotate your credentials

## Requirements

- Node.js 18+
- Google API credentials
- Gmail API access
- Google Sheets API access
- Google Drive API access