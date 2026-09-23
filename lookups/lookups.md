# Excel Lookup Functions

A quick reference guide to three commonly used Excel lookup functions: **XLOOKUP, VLOOKUP, and HLOOKUP**.

These functions are used to **find a value in a table and return related information**.

---

## Quick Comparison

| Feature | XLOOKUP | VLOOKUP | HLOOKUP |
|---|---|---|---|
| Lookup direction | Vertical / Horizontal | Vertical | Horizontal |
| Lookup value location | Any position | First column | First row |
| Return value location | Any position | To the right | Below |
| Exact match | Default | Must specify | Must specify |
| Approximate match | Supported | Supported | Supported |
| Can look left? | Yes | No | N/A |
| Can look right? | Yes | Yes | N/A |
| Can search horizontally? | Yes | No | Yes |
| Newer Excel function | Yes | Legacy | Legacy |
| Recommended for new work | Yes | Sometimes | Sometimes |

---

# 1. XLOOKUP

**XLOOKUP** searches for a value in one range and returns the corresponding value from another range.

### Syntax

```excel
=XLOOKUP(lookup_value, lookup_array, return_array)
```

### Example

Suppose we have:

| Employee ID | Employee | Department | Salary |
|---|---|---|---:|
| 101 | Alice | Sales | $60,000 |
| 102 | Bob | IT | $75,000 |
| 103 | Carol | Finance | $70,000 |

To find Bob's department:

```excel
=XLOOKUP(102, A2:A4, C2:C4)
```

**Result:**

```text
IT
```

### Why use XLOOKUP?

XLOOKUP is generally more flexible than VLOOKUP because it can:

- Search left or right
- Search vertically or horizontally
- Return an exact match by default
- Use separate lookup and return ranges
- Return a custom value when nothing is found

Example with a custom error message:

```excel
=XLOOKUP(105, A2:A4, C2:C4, "Employee Not Found")
```

Result:

```text
Employee Not Found
```

---

# 2. VLOOKUP

**VLOOKUP** searches for a value in the **first column of a table** and returns a value from a column to the right.

### Syntax

```excel
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

### Example

Using the same employee table:

| Employee ID | Employee | Department | Salary |
|---|---|---|---:|
| 101 | Alice | Sales | $60,000 |
| 102 | Bob | IT | $75,000 |
| 103 | Carol | Finance | $70,000 |

To find Bob's department:

```excel
=VLOOKUP(102, A2:D4, 3, FALSE)
```

**Result:**

```text
IT
```

The `3` means:

> Return the value from the 3rd column of the selected table.

`FALSE` means:

> Find an exact match.

### Important limitation

VLOOKUP can only search in the **first column** and return values **to the right**.

For example:

```text
Employee ID → Employee → Department → Salary
```

VLOOKUP can search by Employee ID and return Department.

But it cannot search by Salary and return Employee because Salary is to the right of Employee.

---

# 3. HLOOKUP

**HLOOKUP** works similarly to VLOOKUP, but searches **horizontally across the first row**.

### Syntax

```excel
=HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])
```

### Example

Suppose our data is arranged horizontally:

|  | Alice | Bob | Carol |
|---|---|---|---|
| Employee ID | 101 | 102 | 103 |
| Department | Sales | IT | Finance |
| Salary | $60,000 | $75,000 | $70,000 |

To find Bob's salary:

```excel
=HLOOKUP("Bob", B1:D4, 4, FALSE)
```

**Result:**

```text
$75,000
```

The `4` tells Excel to return the value from the **4th row** of the selected range.

---

# 🔎 XLOOKUP vs. VLOOKUP

### VLOOKUP

```excel
=VLOOKUP(102, A2:D4, 3, FALSE)
```

Excel searches:

```text
102
↓
Employee ID
↓
Bob
↓
IT
```

### XLOOKUP

```excel
=XLOOKUP(102, A2:A4, C2:C4)
```

Excel searches:

```text
102
↓
A2:A4
↓
Returns corresponding value
↓
C2:C4
↓
IT
```

The major difference is that **XLOOKUP separates the lookup range from the return range**.

---

# When to Use Each

| Situation | Recommended Function |
|---|---|
| New Excel project | **XLOOKUP** |
| Need to search left | **XLOOKUP** |
| Need exact matching | **XLOOKUP** |
| Working with older Excel | **VLOOKUP** |
| Existing spreadsheet already uses VLOOKUP | **VLOOKUP** |
| Data is arranged horizontally | **HLOOKUP** |
| Need a simple legacy lookup | **VLOOKUP / HLOOKUP** |

---

# Simple Way to Remember

### XLOOKUP

> **"Find this → return that."**

```excel
=XLOOKUP(what_to_find, where_to_find_it, what_to_return)
```

### VLOOKUP

> **"Find this in a column → move right."**

```excel
=VLOOKUP(what_to_find, table, column_number, FALSE)
```

### HLOOKUP

> **"Find this in a row → move down."**

```excel
=HLOOKUP(what_to_find, table, row_number, FALSE)
```

---

## Key Takeaway

**XLOOKUP is the most flexible of the three.**

```text
                 LOOKUP
                    │
       ┌────────────┼────────────┐
       ↓            ↓            ↓
   XLOOKUP      VLOOKUP      HLOOKUP
   Flexible      Vertical    Horizontal
   Modern        Legacy      Legacy
```

For new Excel work, **XLOOKUP is generally the function to learn first**, while understanding VLOOKUP and HLOOKUP is still useful because they appear frequently in existing spreadsheets.
