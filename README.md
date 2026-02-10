# SpendTab

A personal finance tracking add-on for Google Sheets. Track your spending, categorize transactions, and visualize your expenses and income — all within a familiar spreadsheet interface.

## Features

- **Transaction Entry** — Paste transactions directly from your bank's website or enter manually
- **Duplicate Detection** — Automatically identifies and skips duplicate transactions
- **Category Management** — Organize spending into categories and subcategories
- **Monthly Breakdown** — View expenses and income by month with automatic totals
- **Year-over-Year Tracking** — Archive previous years and start fresh each January
- **Input Validation** — Validates dates and amounts before saving

## Installation

1. Open a new or existing Google Sheet
2. Go to **Extensions** → **Apps Script**
3. Copy the contents of `Code.gs` into the script editor
4. Create a new HTML file named `EnterTrans.html` and paste its contents
5. Save and refresh your spreadsheet
6. Click **SpendTab** → **Setup SpendTab** from the menu

## Usage

### Setting Up
1. Click **SpendTab** → **Setup SpendTab** to create the required sheets
2. The add-on will create: Transactions, Expenses, Income, Summary, and Categories sheets

### Entering Transactions
1. Click **SpendTab** → **Enter Transactions**
2. Select your account type (Checking or Credit Card)
3. Paste transaction data from your bank or enter manually
4. Enable "Detect Duplicates" to avoid re-entering existing transactions
5. Click **Submit**

### Managing Categories
- Edit the Expenses and Income sheets to add/remove categories and subcategories
- Run **SpendTab** → **Fix Category Dropdowns** after making changes

## Sheets Overview

| Sheet | Purpose |
|-------|---------|
| Transactions | Raw transaction log with date, description, amount, category |
| Expenses | Monthly expense totals by category/subcategory |
| Income | Monthly income totals by category/subcategory |
| Summary | Dashboard with yearly totals and net income |
| Categories Dropdown | Hidden sheet for dropdown validation |

## Contributing

Contributions are welcome! Please open an issue or submit a pull request.

## License

This project is open source. Originally forked from the "Budget Tracker" Google Apps Script (Thanks Tad Smith for the initial contribution).
