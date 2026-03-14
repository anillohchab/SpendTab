# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

SpendTab is a Google Apps Script add-on for Google Sheets that provides personal finance tracking. There is no build system, package manager, or test framework — the two source files are deployed directly to Google Apps Script.

## Deployment

Copy `Code.gs` and `EnterTrans.html` into a Google Apps Script project bound to a Google Sheet. There is no `.clasp.json` or `appsscript.json` in the repo; deployment is manual via the Apps Script editor.

## Architecture

**Code.gs** — Server-side Google Apps Script backend:
- Global constants (top of file) define column mappings, sheet names, and menu labels
- `onOpen`/`onInstall` create the "SpendTab" menu; `enableBudgetTracker` runs first-time setup
- `setupBudgetSheets` copies sheets from an external template spreadsheet (`TEMPLATE_SHEET_ID`) and wires up named ranges and formulas
- `onEdit` trigger dynamically sets subcategory dropdown validation when a category cell changes in the Transactions sheet
- `fixSheetFormulas` / `fixCategoryAndSubCategory` regenerate SUMIFS, totals, and averages across Expenses/Income sheets — these are the formula-repair entry points
- `enterTransactions` receives data from the dialog, validates entries, deduplicates, and batch-writes to the Transactions sheet
- `createCurrentYearSheets` archives the previous year's sheets and creates new ones for the current year
- `formatString(template, ...)` is a custom utility replacing `{0}`, `{1}`, etc. — used throughout for formula generation

**EnterTrans.html** — Client-side modal dialog (jQuery 3.7.1):
- Supports multiple input formats (Spreadsheet, Chase Bank Online) with dynamic column reordering
- Handles clipboard paste events for bulk transaction entry (both plain text and HTML table parsing)
- Calls `google.script.run.enterTransactions()` on submit

**Data model** — Google Sheets acts as both UI and database:
- `Transactions` sheet: raw transaction log (Type, Date, Post Date, Description, Amount, Category, Subcategory)
- `[Year] Expenses` / `[Year] Income`: category/subcategory breakdowns with monthly SUMIFS formulas referencing the Transactions sheet
- `[Year] Summary`: aggregates from Expenses and Income sheets
- `Categories Dropdown`: hidden sheet driving category/subcategory data validation
- Year-based naming convention (e.g., "2026 Expenses") enables multi-year archival

## Key Patterns

- Column references use `colNameToNum`/`colNumToName` utilities (supports multi-character columns like AA)
- Sheet formulas are generated programmatically via `formatString` with positional placeholders
- Document properties (`PropertiesService`) store the budget tracker version and onEdit trigger ID
- The `onEdit` trigger is installed once and tracked by unique ID to prevent duplicates
