# Excel Formulas

A practical reference guide to commonly used Excel formulas for **data analysis, reporting, business analysis, and everyday spreadsheet work**.

---

## Formula Categories

| Category                       | What It Helps With                    | Examples                                   |
| ------------------------------ | ------------------------------------- | ------------------------------------------ |
| **Basic Calculations**      | Arithmetic and metrics                | `SUM`, `AVERAGE`, `MIN`, `MAX`             |
| **Conditional Logic**       | Making decisions based on conditions  | `IF`, `IFS`, `AND`, `OR`                   |
| **Conditional Aggregation** | Summarizing filtered data             | `SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTIFS`   |
| **Lookups**                 | Finding related information           | `XLOOKUP`, `VLOOKUP`, `HLOOKUP`            |
| **Date & Time**             | Working with dates                    | `TODAY`, `YEAR`, `MONTH`, `DATEDIF`        |
| **Text**                    | Cleaning and manipulating text        | `LEFT`, `RIGHT`, `MID`, `TRIM`, `TEXTJOIN` |
| **Data Cleaning**           | Handling missing/duplicate data       | `TRIM`, `CLEAN`, `UNIQUE`, `FILTER`        |
| **Analysis**                | Measuring relationships and variation | `ROUND`, `MEDIAN`, `STDEV`, `RANK`         |

---

# 1. Basic Calculation Formulas

## `SUM`

Adds values together.

### Syntax

```excel
=SUM(number1, number2, ...)
```

### Example

| Product |  Sales |
| ------- | -----: |
| iPhone  | $1,200 |
| Mac     | $2,500 |
| iPad    |   $800 |

```excel
=SUM(B2:B4)
```

**Result:**

```text
$4,500
```

**Use it for:** total sales, total units, total costs, total transactions.

---

## `AVERAGE`

Calculates the arithmetic mean.

```excel
=AVERAGE(B2:B10)
```

Example:

```text
80 + 90 + 70
───────────── = 80
      3
```

**Result: `80`**

**Use it for:** average sales, average order value, average score, average KPI.

---

## `MIN & MAX`

Find the smallest or largest value.

```excel
=MIN(B2:B10)
=MAX(B2:B10)
```

Example:

| Store |   Sales |
| ----- | ------: |
| A     | $10,000 |
| B     | $15,000 |
| C     |  $8,000 |

```excel
=MIN(B2:B4)
```

→ `$8,000`

```excel
=MAX(B2:B4)
```

→ `$15,000`

---

# 2. Conditional Logic

## `IF`

Returns one result when a condition is true and another when it is false.

### Syntax

```excel
=IF(condition, value_if_true, value_if_false)
```

### Example

```excel
=IF(B2>=100,"Target Met","Below Target")
```

| Sales | Result       |
| ----: | ------------ |
|   125 | Target Met   |
|    85 | Below Target |

### Think of it as:

```text
IF sales ≥ 100
       ↓
   Target Met
       │
       └── Otherwise → Below Target
```

---

## `IFS`

Tests multiple conditions.

```excel
=IFS(
B2>=90,"Excellent",
B2>=75,"Good",
B2>=60,"Average",
B2<60,"Needs Improvement"
)
```

Example:

| Score | Result            |
| ----: | ----------------- |
|    95 | Excellent         |
|    82 | Good              |
|    68 | Average           |
|    45 | Needs Improvement |

**Use `IFS` when you have multiple categories instead of nesting several `IF` statements.**

---

## `AND`

Checks whether **all conditions** are TRUE.

```excel
=AND(B2>=80,C2="Yes")
```

Returns:

```text
TRUE
```

only if both conditions are satisfied.

### Example

```excel
=IF(AND(B2>=80,C2="Yes"),"Qualified","Not Qualified")
```

---

## `OR`

Checks whether **at least one condition** is TRUE.

```excel
=OR(B2>=80,C2="Yes")
```

Returns `TRUE` if either condition is satisfied.

---

# 3. Conditional Aggregation

These formulas are especially useful for **business and data analysis**.

---

## `SUMIFS`

Adds values that meet **multiple conditions**.

### Example

```excel
=SUMIFS(
C2:C100,
A2:A100,"West",
B2:B100,"Mac"
)
```

Meaning:

> Add Sales where Region = West **AND** Product = Mac.

### Mental model

```text
             SALES
               ↑
        ┌──────┴──────┐
        │             │
     Region          Product
      = West          = Mac
        │             │
        └──────┬──────┘
               ↓
          SUM RESULTS
```

---

## `COUNTIF`

Counts cells meeting **one condition**.

```excel
=COUNTIF(A2:A100,"West")
```

Example:

```text
West
West
East
West
South
```

Result:

```text
3
```

---

## `COUNTIFS`

Counts rows meeting **multiple conditions**.

```excel
=COUNTIFS(
A2:A100,"West",
B2:B100,"Mac"
)
```

Meaning:

> Count transactions where Region = West **AND** Product = Mac.

---

# 4. Lookup Formulas

Lookup formulas allow you to **connect information between different parts of a dataset**.

For a detailed guide, see the [Lookup Functions](#) section.

### `XLOOKUP`

```excel
=XLOOKUP(E2,A2:A100,C2:C100)
```

> Find the value in `E2` inside column A and return the corresponding value from column C.

### `VLOOKUP`

```excel
=VLOOKUP(E2,A2:C100,3,FALSE)
```

> Find `E2` in the first column and return the value from column 3.

### `HLOOKUP`

```excel
=HLOOKUP(E2,A1:F3,3,FALSE)
```

> Find `E2` across the first row and return the corresponding value from row 3.

---

# 5. Date & Time Formulas

Dates are extremely common in **business analytics, sales analysis, and reporting**.

---

## `TODAY`

Returns today's date.

```excel
=TODAY()
```

Example:

```text
09/25/2026
```

Useful for calculating how old a record is or whether something is overdue.

---

## `YEAR`, `MONTH`, `DAY`

Extract individual parts of a date.

```excel
=YEAR(A2)
=MONTH(A2)
=DAY(A2)
```

If:

```text
A2 = 09/25/2026
```

Results:

```text
YEAR  → 2026
MONTH → 9
DAY   → 25
```

---

## `DATEDIF`

Calculates the difference between two dates.

```excel
=DATEDIF(A2,B2,"Y")
```

Returns the number of complete years between the two dates.

Other units:

| Code  | Meaning         |
| ----- | --------------- |
| `"Y"` | Complete years  |
| `"M"` | Complete months |
| `"D"` | Total days      |

---

## `EOMONTH`

Returns the last day of a month.

```excel
=EOMONTH(A2,0)
```

If:

```text
A2 = 09/15/2026
```

Result:

```text
09/30/2026
```

💡 Useful for **monthly reporting and financial analysis**.

---

# 6. Text Formulas

Text functions are useful for **cleaning datasets and preparing data for analysis**.

---

## `LEFT`

Extracts characters from the beginning of text.

```excel
=LEFT(A2,3)
```

Example:

```text
A2 = ABC12345
```

Result:

```text
ABC
```

---

## `RIGHT`

Extracts characters from the end.

```excel
=RIGHT(A2,4)
```

Example:

```text
ABC12345
```

Result:

```text
2345
```

---

## `MID`

Extracts characters from the middle.

```excel
=MID(A2,4,3)
```

Meaning:

```text
Start at character 4
      ↓
ABC12345
   └───┘
     123
```

Result:

```text
123
```

---

## `TRIM`

Removes unnecessary spaces.

```excel
=TRIM(A2)
```

Example:

```text
"   John   Smith   "
```

becomes:

```text
"John Smith"
```

💡 **Very useful when cleaning imported datasets.**

---

## `CONCAT`

Combines text.

```excel
=CONCAT(A2," ",B2)
```

Example:

| First Name | Last Name |
| ---------- | --------- |
| John       | Smith     |

```excel
=CONCAT(A2," ",B2)
```

Result:

```text
John Smith
```

---

## `TEXTJOIN`

Combines multiple pieces of text using a delimiter.

```excel
=TEXTJOIN(", ",TRUE,A2:C2)
```

Example:

```text
Apple | Mac | California
```

becomes:

```text
Apple, Mac, California
```

---

---
---
---
---

# 🧮 Key Excel Formulas

A practical reference guide to commonly used Excel formulas for **data analysis, reporting, business analysis, and everyday spreadsheet work**.

---

## 📚 Formula Categories

| Category                       | What It Helps With                    | Examples                                   |
| ------------------------------ | ------------------------------------- | ------------------------------------------ |
| 🔢 **Basic Calculations**      | Arithmetic and metrics                | `SUM`, `AVERAGE`, `MIN`, `MAX`             |
| 🔎 **Conditional Logic**       | Making decisions based on conditions  | `IF`, `IFS`, `AND`, `OR`                   |
| 📊 **Conditional Aggregation** | Summarizing filtered data             | `SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTIFS`   |
| 🔍 **Lookups**                 | Finding related information           | `XLOOKUP`, `VLOOKUP`, `HLOOKUP`            |
| 📅 **Date & Time**             | Working with dates                    | `TODAY`, `YEAR`, `MONTH`, `DATEDIF`        |
| 🔤 **Text**                    | Cleaning and manipulating text        | `LEFT`, `RIGHT`, `MID`, `TRIM`, `TEXTJOIN` |
| 🧹 **Data Cleaning**           | Handling missing/duplicate data       | `TRIM`, `CLEAN`, `UNIQUE`, `FILTER`        |
| 📈 **Analysis**                | Measuring relationships and variation | `ROUND`, `MEDIAN`, `STDEV`, `RANK`         |

---

# 🔢 1. Basic Calculation Formulas

## `SUM`

Adds values together.

### Syntax

```excel
=SUM(number1, number2, ...)
```

### Example

| Product |  Sales |
| ------- | -----: |
| iPhone  | $1,200 |
| Mac     | $2,500 |
| iPad    |   $800 |

```excel
=SUM(B2:B4)
```

**Result:**

```text
$4,500
```

💡 **Use it for:** total sales, total units, total costs, total transactions.

---

## `AVERAGE`

Calculates the arithmetic mean.

```excel
=AVERAGE(B2:B10)
```

Example:

```text
80 + 90 + 70
───────────── = 80
      3
```

**Result: `80`**

💡 **Use it for:** average sales, average order value, average score, average KPI.

---

## `MIN & MAX`

Find the smallest or largest value.

```excel
=MIN(B2:B10)
=MAX(B2:B10)
```

Example:

| Store |   Sales |
| ----- | ------: |
| A     | $10,000 |
| B     | $15,000 |
| C     |  $8,000 |

```excel
=MIN(B2:B4)
```

→ `$8,000`

```excel
=MAX(B2:B4)
```

→ `$15,000`

---

# 🧠 2. Conditional Logic

## `IF`

Returns one result when a condition is true and another when it is false.

### Syntax

```excel
=IF(condition, value_if_true, value_if_false)
```

### Example

```excel
=IF(B2>=100,"Target Met","Below Target")
```

| Sales | Result       |
| ----: | ------------ |
|   125 | Target Met   |
|    85 | Below Target |

### Think of it as:

```text
IF sales ≥ 100
       ↓
   Target Met
       │
       └── Otherwise → Below Target
```

---

## `IFS`

Tests multiple conditions.

```excel
=IFS(
B2>=90,"Excellent",
B2>=75,"Good",
B2>=60,"Average",
B2<60,"Needs Improvement"
)
```

Example:

| Score | Result            |
| ----: | ----------------- |
|    95 | Excellent         |
|    82 | Good              |
|    68 | Average           |
|    45 | Needs Improvement |

💡 **Use `IFS` when you have multiple categories instead of nesting several `IF` statements.**

---

## `AND`

Checks whether **all conditions** are TRUE.

```excel
=AND(B2>=80,C2="Yes")
```

Returns:

```text
TRUE
```

only if both conditions are satisfied.

### Example

```excel
=IF(AND(B2>=80,C2="Yes"),"Qualified","Not Qualified")
```

---

## `OR`

Checks whether **at least one condition** is TRUE.

```excel
=OR(B2>=80,C2="Yes")
```

Returns `TRUE` if either condition is satisfied.

---


# 🔎 4. Lookup Formulas

Lookup formulas allow you to **connect information between different parts of a dataset**.

For a detailed guide, see the [Lookup Functions](#) section.

### `XLOOKUP`

```excel
=XLOOKUP(E2,A2:A100,C2:C100)
```

> Find the value in `E2` inside column A and return the corresponding value from column C.

### `VLOOKUP`

```excel
=VLOOKUP(E2,A2:C100,3,FALSE)
```

> Find `E2` in the first column and return the value from column 3.

### `HLOOKUP`

```excel
=HLOOKUP(E2,A1:F3,3,FALSE)
```

> Find `E2` across the first row and return the corresponding value from row 3.

---

# 📅 5. Date & Time Formulas

Dates are extremely common in **business analytics, sales analysis, and reporting**.

---

## `TODAY`

Returns today's date.

```excel
=TODAY()
```

Example:

```text
09/25/2026
```

Useful for calculating how old a record is or whether something is overdue.

---

## `YEAR`, `MONTH`, `DAY`

Extract individual parts of a date.

```excel
=YEAR(A2)
=MONTH(A2)
=DAY(A2)
```

If:

```text
A2 = 09/25/2026
```

Results:

```text
YEAR  → 2026
MONTH → 9
DAY   → 25
```

---

## `DATEDIF`

Calculates the difference between two dates.

```excel
=DATEDIF(A2,B2,"Y")
```

Returns the number of complete years between the two dates.

Other units:

| Code  | Meaning         |
| ----- | --------------- |
| `"Y"` | Complete years  |
| `"M"` | Complete months |
| `"D"` | Total days      |

---

## `EOMONTH`

Returns the last day of a month.

```excel
=EOMONTH(A2,0)
```

If:

```text
A2 = 09/15/2026
```

Result:

```text
09/30/2026
```

💡 Useful for **monthly reporting and financial analysis**.

---

# 🔤 6. Text Formulas

Text functions are useful for **cleaning datasets and preparing data for analysis**.

---

## `LEFT`

Extracts characters from the beginning of text.

```excel
=LEFT(A2,3)
```

Example:

```text
A2 = ABC12345
```

Result:

```text
ABC
```

---

## `RIGHT`

Extracts characters from the end.

```excel
=RIGHT(A2,4)
```

Example:

```text
ABC12345
```

Result:

```text
2345
```

---

## `MID`

Extracts characters from the middle.

```excel
=MID(A2,4,3)
```

Meaning:

```text
Start at character 4
      ↓
ABC12345
   └───┘
     123
```

Result:

```text
123
```

---

## `TRIM`

Removes unnecessary spaces.

```excel
=TRIM(A2)
```

Example:

```text
"   John   Smith   "
```

becomes:

```text
"John Smith"
```

💡 **Very useful when cleaning imported datasets.**

---

## `CONCAT`

Combines text.

```excel
=CONCAT(A2," ",B2)
```

Example:

| First Name | Last Name |
| ---------- | --------- |
| John       | Smith     |

```excel
=CONCAT(A2," ",B2)
```

Result:

```text
John Smith
```

---

## `TEXTJOIN`

Combines multiple pieces of text using a delimiter.

```excel
=TEXTJOIN(", ",TRUE,A2:C2)
```

Example:

```text
Apple | Mac | California
```

becomes:

```text
Apple, Mac, California
```

---

# 🧹 7. Data Cleaning & Dynamic Arrays

Modern Excel includes functions that make data preparation much easier.

---

## `UNIQUE`

Returns unique values from a range.

```excel
=UNIQUE(A2:A100)
```

Example:

```text
West
East
West
South
East
```

becomes:

```text
West
East
South
```

---

## `FILTER`

Returns only rows that meet a condition.

```excel
=FILTER(A2:C100,C2:C100="West")
```

Meaning:

> Return rows from A:C where the Region column equals West.

### Mental model

```text
FULL DATASET
     │
     ↓
  FILTER
     │
     ├── Region = West
     │
     ↓
WEST REGION DATA
```

---

## `SORT`

Sorts a range dynamically.

```excel
=SORT(A2:B100,2,-1)
```

Meaning:

> Sort the dataset by column 2 in descending order.

---

# 📈 8. Statistical & Analysis Formulas

## `MEDIAN`

Returns the middle value.

```excel
=MEDIAN(B2:B10)
```

Useful when extreme values could distort an average.

---

## `STDEV.S`

Calculates sample standard deviation.

```excel
=STDEV.S(B2:B100)
```

A larger standard deviation generally indicates that the values are more spread out around the mean.

---

## `RANK`

Ranks a value relative to other values.

```excel
=RANK(B2,$B$2:$B$10,0)
```

Example:

| Employee |   Sales | Rank |
| -------- | ------: | ---: |
| Alice    | $10,000 |    2 |
| Bob      | $12,000 |    1 |
| Carol    |  $8,000 |    3 |

---

# 🧩 Formula Combinations

The real power of Excel comes from **combining formulas**.

### Example: IF + XLOOKUP

```excel
=IF(
XLOOKUP(E2,A2:A100,C2:C100)>=100,
"Target Met",
"Below Target"
)
```

This combines:

```text
XLOOKUP
   ↓
Find the employee's sales
   ↓
IF
   ↓
Compare sales against target
   ↓
Return category
```

---

### Example: SUMIFS + DATE

```excel
=SUMIFS(
C2:C100,
A2:A100,">="&DATE(2026,1,1),
A2:A100,"<="&DATE(2026,1,31)
)
```

This calculates:

> Total sales occurring during January 2026.

---

# 🧠 Quick Formula Cheat Sheet

| Formula    | Primary Purpose             | Example                             |
| ---------- | --------------------------- | ----------------------------------- |
| `SUM`      | Add values                  | `=SUM(B2:B10)`                      |
| `AVERAGE`  | Calculate mean              | `=AVERAGE(B2:B10)`                  |
| `MIN`      | Find smallest value         | `=MIN(B2:B10)`                      |
| `MAX`      | Find largest value          | `=MAX(B2:B10)`                      |
| `IF`       | Conditional logic           | `=IF(B2>100,"Yes","No")`            |
| `IFS`      | Multiple conditions         | `=IFS(B2>=90,"A",B2>=80,"B")`       |
| `AND`      | All conditions must be true | `=AND(B2>50,C2="Yes")`              |
| `OR`       | At least one condition      | `=OR(B2>50,C2="Yes")`               |
| `SUMIF`    | Conditional sum             | `=SUMIF(A:A,"West",B:B)`            |
| `SUMIFS`   | Multiple-condition sum      | `=SUMIFS(C:C,A:A,"West",B:B,"Mac")` |
| `COUNTIF`  | Conditional count           | `=COUNTIF(A:A,"West")`              |
| `COUNTIFS` | Multiple-condition count    | `=COUNTIFS(A:A,"West",B:B,"Mac")`   |
| `XLOOKUP`  | Flexible lookup             | `=XLOOKUP(E2,A:A,C:C)`              |
| `VLOOKUP`  | Vertical lookup             | `=VLOOKUP(E2,A:C,3,FALSE)`          |
| `HLOOKUP`  | Horizontal lookup           | `=HLOOKUP(E2,A1:F3,3,FALSE)`        |
| `TODAY`    | Current date                | `=TODAY()`                          |
| `YEAR`     | Extract year                | `=YEAR(A2)`                         |
| `MONTH`    | Extract month               | `=MONTH(A2)`                        |
| `DATEDIF`  | Date difference             | `=DATEDIF(A2,B2,"D")`               |
| `LEFT`     | Extract left characters     | `=LEFT(A2,3)`                       |
| `RIGHT`    | Extract right characters    | `=RIGHT(A2,4)`                      |
| `MID`      | Extract middle characters   | `=MID(A2,4,3)`                      |
| `TRIM`     | Remove extra spaces         | `=TRIM(A2)`                         |
| `UNIQUE`   | Return unique values        | `=UNIQUE(A:A)`                      |
| `FILTER`   | Filter data dynamically     | `=FILTER(A:C,C:C="West")`           |
| `SORT`     | Sort data dynamically       | `=SORT(A:B,2,-1)`                   |
| `MEDIAN`   | Find middle value           | `=MEDIAN(B:B)`                      |
| `STDEV.S`  | Measure variation           | `=STDEV.S(B:B)`                     |
| `RANK`     | Rank values                 | `=RANK(B2,B:B,0)`                   |

---

# 🎯 What to Learn First

For **business and data analytics**, a practical learning progression is:

```text
FOUNDATION
   │
   ├── SUM
   ├── AVERAGE
   └── MIN / MAX
        │
        ↓
CONDITIONAL ANALYSIS
   │
   ├── IF
   ├── AND / OR
   ├── SUMIFS
   └── COUNTIFS
        │
        ↓
DATA LOOKUPS
   │
   └── XLOOKUP
        │
        ↓
DATA CLEANING
   │
   ├── TRIM
   ├── LEFT / RIGHT / MID
   ├── UNIQUE
   └── FILTER
        │
        ↓
ANALYSIS
   │
   ├── MEDIAN
   ├── STDEV.S
   └── RANK
        │
        ↓
COMBINE FORMULAS
   │
   └── Build real-world analysis
```

> **Key idea:** Don't just memorize formulas. Learn **what problem each formula solves**, then learn how to combine them.


