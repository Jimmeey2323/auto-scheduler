# Sold Out Feature

## Overview

Classes can be marked as "Sold Out" which will display a red strikethrough line across the class row and a red "SOLD OUT" badge.

## How to Use

### Method 1: Set Theme to "Sold Out" in Google Sheets

In the Google Sheets "Cleaned" tab, set the **Theme** column (column H) to exactly `Sold Out` for any class you want to mark as sold out.

**Example:**
```
| Day    | Time     | Location      | Class      | Trainer  | Notes | Date        | Theme     |
|--------|----------|---------------|------------|----------|-------|-------------|-----------|
| Monday | 07:15 AM | Kwality House | Barre 57   | Anisha   |       | 09 Feb 2026 | Sold Out  |
```

### Method 2: Add "SOLD OUT" in Notes Column

Alternatively, include `SOLD OUT` anywhere in the **Notes** column (column F).

**Example:**
```
| Day    | Time     | Location      | Class      | Trainer  | Notes     | Date        | Theme |
|--------|----------|---------------|------------|----------|-----------|-------------|-------|
| Monday | 07:15 AM | Kwality House | Barre 57   | Anisha   | SOLD OUT  | 09 Feb 2026 |       |
```

## What It Does

When a class is marked as sold out:

1. ✅ **Red Strikethrough Line** - A crimson (#dc143c) line is drawn across the time and class name
2. ✅ **"SOLD OUT" Badge** - A red badge with white text appears next to the class name
3. ✅ **Theme Badge Suppressed** - If theme is set to "Sold Out", no separate theme badge is shown (only the sold out badge)
4. ✅ **Existing Themes Preserved** - If using Notes method, any existing theme badge will still be shown alongside the sold out badge

## Visual Example

**Normal class:**
```
07:15 AM    BARRE 57 - ANISHA
```

**Sold out class:**
```
07:15 AM    BARRE 57 - ANISHA [SOLD OUT]
───────────────────────────────
(red strikethrough line across entire row)
```

## Running the Update

After setting classes as sold out in the Cleaned sheet:

```bash
# Full workflow (processes email + updates HTML)
npm run update

# HTML-only mode (uses existing Cleaned sheet data)
npm run update:skip-email
```

## Technical Details

- **Strikethrough line width:** 220px (covers time and class name, but not badges)
- **Line color:** #dc143c (crimson red)
- **Line height:** 2px
- **Badge color:** #dc143c background, white text
- **Badge styling:** Rounded corners (4px), bold font (700 weight)
- **Z-index handling:** Strikethrough at z-index 10, badge at z-index 20

## Case Insensitive

The theme check is case-insensitive, so these all work:
- `Sold Out`
- `sold out`
- `SOLD OUT`
- `SoLd OuT`
