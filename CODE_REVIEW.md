# Code Review: SpendTab

## Overview
Review of the SpendTab personal finance tracking Google Apps Script codebase covering `Code.gs` and `EnterTrans.html`.

---

## Issues

### High Priority

- [x] **Implicit Global Variables** (`Code.gs:2-44`)
  All top-level constants are declared without `var`/`const`, polluting the global scope.

- [x] **XSS in Duplicate Transaction Dialog** (`Code.gs:614-626`)
  Transaction descriptions are inserted into HTML without escaping. Malicious descriptions could inject scripts.

- [x] **Invalid CSS `style` Attribute Syntax** (`EnterTrans.html:63,74`)
  Uses `width=100%` and `justify-content=center` instead of `width:100%` and `justify-content:center`.

### Medium Priority

- [x] **Typo `budgeTrackerVersion`** (`Code.gs:66,83`)
  Property key is missing a "t" -- should be `budgetTrackerVersion`.

- [x] **`catch(e)` Shadows Event Parameter** (`Code.gs:153`)
  The catch variable `e` shadows the outer `onEdit(e)` event parameter.

- [x] **No Input Validation on Transactions** (`Code.gs:568+`)
  `enterTransactions` does not validate date format, numeric amounts, or required fields.

- [x] **Inefficient Cell-by-Cell Writes** (`Code.gs:605-610`)
  Each transaction field is written with a separate `setValue()` call. Should batch with `setValues()`.

- [x] **Invalid CSS `//` Comments** (`EnterTrans.html:13,52`)
  Uses `//` which is not valid CSS comment syntax. Should be `/* */`.

- [x] **Function Typo `handePasteEvent`** (`EnterTrans.html:109,255`)
  Missing "l" -- should be `handlePasteEvent`.

- [x] **Variable Shadowing `row`** (`EnterTrans.html:401-402`)
  Callback parameter `row` is immediately shadowed by `var row = $(row)`.

### Low Priority

- [ ] **`String.prototype` Pollution** (`Code.gs:684-706`)
  Monkey-patches `format()`, `startsWith()`, `endsWith()` onto String prototype. Polyfills are unnecessary in V8 runtime.

- [ ] **Missing Semicolons** (`Code.gs` various lines)
  Several statements lack trailing semicolons.

- [ ] **Inconsistent Null Checks** (`Code.gs` throughout)
  Mixes `== null` (loose) and `!== null` (strict) comparison.

- [ ] **Magic Numbers** (`Code.gs:345,378`)
  Unexplained hardcoded values like 60, 25, 20 in range calls. `NUM_ROWS` is declared but unused.

- [ ] **Dead Code and TODOs** (both files)
  Commented-out `Logger.log` calls, unresolved `FIXME`/`TODO` blocks.

- [ ] **Outdated jQuery Version** (`EnterTrans.html:101`)
  Uses jQuery 3.3.1; current versions have security and performance fixes.

- [ ] **Single-Character Column Limit** (`Code.gs:535-541`)
  `colNameToNum`/`colNumToName` only handle columns A-Z.
