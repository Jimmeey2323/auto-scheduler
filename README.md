# Auto Scheduler

An automated system for managing and updating class schedules, integrating with Google Sheets and generating PDF schedules.

## Features

- 📅 Automated schedule updates from Google Sheets
- 🔄 HTML template processing for different locations (Bandra, Kemps)
- 📧 Email notifications with PDF attachments
- 🎨 Theme-based styling (PowerCycle, Total, Core Yoga)
- 📊 Google Sheets integration for real-time data

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

### Update HTML Only
```bash
node updateHTMLOnly.js
```

### Update Kemps Schedule (with email)
```bash
node updateKempsSchedule.js
```

### Complete Schedule Update
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