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

3. Conditional Aggregation

These formulas are especially useful for business and data analysis.

SUMIF

Adds values that meet one condition.

Example
Region	Sales
West	$500
East	$700
West	$300
South	$900
=SUMIF(A2:A5,"West",B2:B5)

Result:

$800

Because:

West → $500
West → $300
      ─────
      $800
SUMIFS

Adds values that meet multiple conditions.

Example
=SUMIFS(
C2:C100,
A2:A100,"West",
B2:B100,"Mac"
)

Meaning:

Add Sales where Region = West AND Product = Mac.

Mental model
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
COUNTIF

Counts cells meeting one condition.

=COUNTIF(A2:A100,"West")

Example:

West
West
East
West
South

Result:

3
COUNTIFS

Counts rows meeting multiple conditions.

=COUNTIFS(
A2:A100,"West",
B2:B100,"Mac"
)

Meaning:

Count transactions where Region = West AND Product = Mac.
