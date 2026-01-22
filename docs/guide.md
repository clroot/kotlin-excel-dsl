# Usage Guide

[한국어](guide.ko.md)

This guide covers styling, advanced features, and detailed usage of the Excel DSL.

## Table of Contents

- [Styling](#styling)
  - [Global Styles](#global-styles)
  - [Column-Level Styles](#column-level-styles)
  - [Inline Column Styles](#inline-column-styles)
  - [Separate Header and Body Styles](#separate-header-and-body-styles)
  - [Style Priority](#style-priority)
  - [Themes](#themes)
- [Advanced Features](#advanced-features)
  - [Header Groups](#header-groups)
  - [Column Width](#column-width)
  - [Freeze Panes](#freeze-panes)
  - [Auto Filter](#auto-filter)
  - [Alternate Row Styling](#alternate-row-styling)
  - [Multi-Sheet](#multi-sheet)

## Styling

### Global Styles

Apply styles to all headers and body cells across the document:

```kotlin
excel {
    styles {
        header {
            backgroundColor(Color.GRAY)
            fontColor(Color.WHITE)
            bold()
            align(Alignment.CENTER)
        }
        body {
            border(BorderStyle.THIN)
        }
    }
    sheet<User>("Users") {
        column("Name") { it.name }
        column("Age") { it.age }
        rows(users)
    }
}.writeTo(output)
```

### Column-Level Styles

Target specific columns using the `styles` DSL:

```kotlin
excel {
    styles {
        header {
            backgroundColor(Color.GRAY)
            bold()
        }
        column("Price") {
            header {
                backgroundColor(Color.BLUE)
            }
            body {
                align(Alignment.RIGHT)
                numberFormat("#,##0")
            }
        }
    }
    sheet<Product>("Products") {
        column("Name") { it.name }
        column("Price") { it.price }
        rows(products)
    }
}.writeTo(output)
```

### Inline Column Styles

Apply styles directly when defining a column:

```kotlin
excel {
    sheet<Product>("Products") {
        column("Name") { it.name }
        column("Price", style = { align(Alignment.RIGHT); bold() }) { it.price }
        rows(products)
    }
}.writeTo(output)
```

### Separate Header and Body Styles

Define different styles for header and body cells of a column:

```kotlin
excel {
    sheet<Product>("Products") {
        column("Name") { it.name }
        column(
            "Price",
            headerStyle = { backgroundColor(Color.BLUE); bold() },
            bodyStyle = { align(Alignment.RIGHT) }
        ) { it.price }
        rows(products)
    }
}.writeTo(output)
```

### Style Priority

When multiple styles are defined, they are applied in this order (highest priority first):

1. **Inline styles** - Defined directly on the column
2. **Column-specific styles** - Defined via `styles { column("Name") { ... } }`
3. **Global styles** - Defined via `styles { header { ... } }`

### Themes

Use predefined themes for consistent styling:

```kotlin
excel(theme = Theme.Modern) {
    sheet<User>("Users") {
        column("Name") { it.name }
        column("Age") { it.age }
        rows(users)
    }
}.writeTo(output)
```

| Theme | Description |
|-------|-------------|
| `Theme.Modern` | Blue header with white text, centered |
| `Theme.Minimal` | Bold header, no background |
| `Theme.Classic` | Gray header, medium borders |

## Advanced Features

### Header Groups

Create multi-row headers with automatic cell merging:

```kotlin
excel {
    sheet<Student>("Report") {
        headerGroup("Student Info") {
            column("Name") { it.name }
            column("ID") { it.studentId }
        }
        headerGroup("Scores") {
            column("Math") { it.math }
            column("English") { it.english }
        }
        rows(students)
    }
}.writeTo(output)
```

Result:

```
+-------Student Info-------+--------Scores--------+
+----Name----+-----ID------+---Math---+--English--+
|   Alice    |   20241     |    95    |    88     |
```

### Column Width

Set fixed column widths using the `chars` extension:

```kotlin
sheet<User>("Users") {
    column("Name", width = 20.chars) { it.name }
    column("Description", width = 50.chars) { it.description }
    rows(users)
}
```

By default, columns use auto-width which calculates the optimal width based on content.

### Freeze Panes

Lock rows and/or columns to keep them visible while scrolling:

```kotlin
excel {
    sheet<User>("Users") {
        freezePane(row = 1)              // Freeze header row
        freezePane(row = 1, col = 1)     // Freeze header row and first column
        column("Name") { it.name }
        column("Age") { it.age }
        rows(users)
    }
}.writeTo(output)
```

Parameters:
- `row` - Number of rows to freeze from the top (default: 0)
- `col` - Number of columns to freeze from the left (default: 0)

### Auto Filter

Add filter dropdowns to the header row:

```kotlin
excel {
    sheet<User>("Users") {
        autoFilter()
        column("Name") { it.name }
        column("Age") { it.age }
        column("Department") { it.department }
        rows(users)
    }
}.writeTo(output)
```

When header groups are present, the filter is applied to the column header row (second row).

### Alternate Row Styling

Apply zebra stripe styling for better readability:

```kotlin
excel {
    sheet<User>("Users") {
        alternateRowStyle {
            backgroundColor(Color.LIGHT_GRAY)
        }
        column("Name") { it.name }
        column("Age") { it.age }
        rows(users)
    }
}.writeTo(output)
```

The style is applied to even-indexed data rows (0, 2, 4, ...) and **merges** with existing body styles. This means you can combine `alternateRowStyle` with `bodyStyle`:

```kotlin
excel {
    styles {
        body {
            bold()
            align(Alignment.CENTER)
        }
    }
    sheet<User>("Users") {
        alternateRowStyle {
            backgroundColor(Color.LIGHT_GRAY)
        }
        // Even rows will have: bold + center + gray background
        // Odd rows will have: bold + center
        column("Name") { it.name }
        rows(users)
    }
}.writeTo(output)
```

### Multi-Sheet

Create workbooks with multiple sheets:

```kotlin
excel {
    sheet<Product>("Products") {
        column("Name") { it.name }
        column("Price") { it.price }
        rows(products)
    }
    sheet<Order>("Orders") {
        column("Order ID") { it.id }
        column("Amount") { it.amount }
        rows(orders)
    }
}.writeTo(output)
```
