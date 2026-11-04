# Usage Guide

[한국어](guide.ko.md)

This guide covers styling, advanced features, and detailed usage of the Excel DSL.

## Table of Contents

- [Styling](#styling)
  - [Global Styles](#global-styles)
  - [Column-Level Styles](#column-level-styles)
  - [Inline Column Styles](#inline-column-styles)
  - [Separate Header and Body Styles](#separate-header-and-body-styles)
  - [Conditional Styling](#conditional-styling)
  - [Number and Date Formats](#number-and-date-formats)
  - [Style Priority](#style-priority)
  - [Themes](#themes)
- [Annotation Styling](#annotation-styling)
  - [HeaderStyle and BodyStyle](#headerstyle-and-bodystyle)
  - [ConditionalStyle](#conditionalstyle)
  - [Style Enums](#style-enums)
  - [HEX Color Support](#hex-color-support)
  - [Annotation Style Priority](#annotation-style-priority)
- [Advanced Features](#advanced-features)
  - [Header Groups](#header-groups)
  - [Column Width](#column-width)
  - [Freeze Panes](#freeze-panes)
  - [Auto Filter](#auto-filter)
  - [Alternate Row Styling](#alternate-row-styling)
  - [Formulas](#formulas)
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

### Conditional Styling

Apply dynamic styles based on cell values at render time:

```kotlin
excel {
    sheet<Transaction>("Transactions") {
        column("Description") { it.description }
        column("Amount", conditionalStyle = { value: Int? ->
            when {
                value == null -> null
                value < 0 -> fontColor(Color.RED)
                value > 1000000 -> fontColor(Color.GREEN)
                else -> null
            }
        }) { it.amount }
        rows(transactions)
    }
}.writeTo(output)
```

Conditional styles can be combined with body styles:

```kotlin
excel {
    sheet<Transaction>("Transactions") {
        column(
            "Amount",
            bodyStyle = { bold(); align(Alignment.RIGHT) },
            conditionalStyle = { value: Int? ->
                if (value != null && value < 0) fontColor(Color.RED) else null
            }
        ) { it.amount }
        // Negative values will be: bold + right-aligned + red
        // Positive values will be: bold + right-aligned
        rows(transactions)
    }
}.writeTo(output)
```

**Convenience functions** for simple cases:

```kotlin
import io.clroot.excel.core.dsl.fontColor
import io.clroot.excel.core.dsl.backgroundColor

column("Status", conditionalStyle = { value: String? ->
    when (value) {
        "ERROR" -> fontColor(Color.RED)
        "SUCCESS" -> fontColor(Color.GREEN)
        else -> null
    }
}) { it.status }
```

### Number and Date Formats

Set a column's body number format with the `format` parameter. It is written to the
cell's Excel number format (`dataFormatString`); header cells are not affected:

```kotlin
excel {
    sheet<Order>("Orders") {
        column("Amount", format = "#,##0") { it.amount }
        column("Ordered", format = "yyyy/mm/dd") { it.orderedAt }
        rows(orders)
    }
}.writeTo(output)
```

Annotations use `@Column(format = "...")` with the same behavior:

```kotlin
@Excel
data class Order(
    @Column("Amount", format = "#,##0")
    val amount: Long,
)
```

Rules:
- `format` takes precedence over any `numberFormat(...)` from global, column-specific
  or inline body styles of the same column. Other style properties are still applied.
- A conditional style that sets `numberFormat` overrides `format` for that cell.
- `LocalDate` and `LocalDateTime` cells use the resolved number format: `format`, or a
  `numberFormat(...)` from any style layer. Only when none is set do they use
  `yyyy-mm-dd` and `yyyy-mm-dd hh:mm:ss`. A global `styles { body { numberFormat(...) } }`
  therefore also applies to date columns; set `format` on date columns to keep a date format.
- Patterns are Excel number format codes (for example `#,##0.00`, `0%`, `yyyy-mm-dd`),
  not Java `DateTimeFormatter` patterns.

### Style Priority

Style layers are **merged**, property by property. A property set in a higher layer wins;
a property a layer leaves unset is inherited from the layers below. No layer discards
the others, so an inline `bold()` keeps the borders and alignment from global styles.

Body cells, highest priority first:

1. **Conditional styles** - Dynamic styles based on cell values
2. **Column format** - `column(..., format = "...")`, applied as the number format
3. **Inline styles** - `column(..., style = { ... })` or `bodyStyle = { ... }`
4. **Column-specific styles** - Defined via `styles { column("Name") { body { ... } } }`
5. **Global styles** - `styles { body { ... } }`, merged over the theme's body style
6. **Alternate row styles** - `alternateRowStyle { ... }`, even-indexed rows only

Header cells follow the same rule: inline `headerStyle = { ... }` >
`styles { column("Name") { header { ... } } }` > `styles { header { ... } }` > theme.
Group header cells use only the global header style.

Bold and italic are tri-state: `bold()` / `bold(true)` turns bold on, `bold(false)` turns it
off, and not calling either inherits it. A higher layer can therefore turn off bold set
by a lower layer, for example with a conditional style:

```kotlin
column(
    "Amount",
    bodyStyle = { bold() },
    conditionalStyle = { value: Int? ->
        if (value == 0) CellStyle(bold = false) else null
    },
) { it.amount }
```

`LocalDate`/`LocalDateTime` cells get the default date format only when no layer sets a
number format (see [Number and Date Formats](#number-and-date-formats)).

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

## Annotation Styling

When using the annotation approach (`@Excel`, `@Column`), you can apply styles using `@HeaderStyle`, `@BodyStyle`, and `@ConditionalStyle` annotations.

### HeaderStyle and BodyStyle

Apply styles at class level (all columns) or property level (specific column):

```kotlin
@Excel
@HeaderStyle(bold = true, backgroundColor = StyleColor.LIGHT_GRAY)
@BodyStyle(alignment = StyleAlignment.CENTER)
data class User(
    @Column("Name", order = 1)
    @HeaderStyle(fontColor = StyleColor.BLUE)  // Override class-level
    val name: String,

    @Column("Age", order = 2)
    @BodyStyle(bold = true)  // Override class-level
    val age: Int,

    @Column("Email", order = 3)
    val email: String  // Uses class-level styles
)

excelOf(users).writeTo(output)
```

Available attributes:
- `bold` - Bold font (default: false)
- `italic` - Italic font (default: false)
- `backgroundColor` - Background color using `StyleColor` enum
- `backgroundColorHex` - Background color using HEX string (e.g., "#FF5733")
- `fontColor` - Font color using `StyleColor` enum
- `fontColorHex` - Font color using HEX string
- `alignment` - Text alignment using `StyleAlignment` enum
- `border` - Border style using `StyleBorder` enum

### ConditionalStyle

Apply dynamic styles based on cell values using the `@ConditionalStyle` annotation and `ConditionalStyler<T>` interface:

```kotlin
class ScoreStyler : ConditionalStyler<Int> {
    override fun style(value: Int?): CellStyle? = when {
        value == null -> null
        value >= 90 -> CellStyle(fontColor = Color.GREEN, bold = true)
        value >= 70 -> CellStyle(fontColor = Color.BLUE)
        value < 60 -> CellStyle(fontColor = Color.RED)
        else -> null
    }
}

class StatusStyler : ConditionalStyler<String> {
    override fun style(value: String?): CellStyle? = when (value) {
        "ACTIVE" -> CellStyle(backgroundColor = Color(200, 255, 200))
        "INACTIVE" -> CellStyle(backgroundColor = Color(255, 200, 200))
        else -> null
    }
}

@Excel
data class Student(
    @Column("Name", order = 1)
    val name: String,

    @Column("Score", order = 2)
    @ConditionalStyle(ScoreStyler::class)
    val score: Int,

    @Column("Status", order = 3)
    @ConditionalStyle(StatusStyler::class)
    val status: String
)
```

**Important:** The generic type of `ConditionalStyler<T>` must match the property type. A type mismatch will throw `ExcelConfigurationException` at runtime.

### Style Enums

Use these enums for type-safe style definitions:

**StyleColor:**
```kotlin
enum class StyleColor {
    NONE,        // No color (inherit)
    WHITE, BLACK, GRAY, LIGHT_GRAY,
    RED, GREEN, BLUE, YELLOW, ORANGE
}
```

**StyleAlignment:**
```kotlin
enum class StyleAlignment {
    NONE,    // No alignment (inherit)
    LEFT, CENTER, RIGHT
}
```

**StyleBorder:**
```kotlin
enum class StyleBorder {
    NONE,    // No border (inherit)
    THIN, MEDIUM, THICK
}
```

### HEX Color Support

For colors not available in `StyleColor`, use HEX strings:

```kotlin
@Excel
@HeaderStyle(backgroundColorHex = "#4A90D9", fontColorHex = "#FFFFFF")
data class CustomStyled(
    @Column("Value", order = 1)
    @BodyStyle(backgroundColorHex = "#F5F5F5")
    val value: String
)
```

Supported formats:
- 6-digit HEX: `"#RRGGBB"` (e.g., "#FF5733")
- 6-digit without #: `"RRGGBB"` (e.g., "FF5733")

**Note:** If both enum color and HEX color are specified, HEX takes precedence.

### Annotation Style Priority

Styles are merged with the same rule as the DSL (higher layers win per property,
unset properties are inherited), in this order (highest priority first):

1. **ConditionalStyle** - Dynamic styles from `ConditionalStyler`
2. **Property-level annotations** - `@HeaderStyle`/`@BodyStyle` on properties
3. **Class-level annotations** - `@HeaderStyle`/`@BodyStyle` on class
4. **Theme styles** - From `excelOf(data, theme = Theme.Modern)`

Example:
```kotlin
@Excel
@HeaderStyle(bold = true)  // Priority 3
@BodyStyle(alignment = StyleAlignment.LEFT)  // Priority 3
data class Report(
    @Column("Score", order = 1)
    @BodyStyle(alignment = StyleAlignment.RIGHT)  // Priority 2 - overrides class
    @ConditionalStyle(ScoreStyler::class)  // Priority 1 - adds conditional
    val score: Int
)

// With theme:
excelOf(reports, theme = Theme.Modern)  // Theme is Priority 4
```

`@Column(format = "...")` sits between property-level annotations and `ConditionalStyle`,
as in the DSL. Annotation attributes cannot express "not specified" for booleans, so
`@HeaderStyle(bold = false)` / `@BodyStyle(bold = false)` (the default) does not turn off
bold from a theme or class-level annotation. Return `CellStyle(bold = false)` from a
`ConditionalStyler` when a cell must not be bold.

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

Auto-width estimates the text each cell displays, not the raw value's `toString()`:

- Numbers are measured in plain notation (`10000000`, not `1.0E7`). When a number format
  applies (`format`, or `numberFormat` from a style), the formatted text is measured,
  including thousands separators and decimals.
- Dates are measured with the applied date format (by default `yyyy-mm-dd` and
  `yyyy-mm-dd hh:mm:ss`).
- Formula results are unknown at write time, so formula cells count as a fixed 10 characters.
  Set a fixed width if the results are wider.
- CJK characters count as two characters. Font size and bold text are not taken into account.
- Widths include 2 characters of padding and are clamped to 8–100 characters.

Only the maximum width per column is kept while streaming rows, so memory use does not
grow with the number of rows.

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

The style is applied to even-indexed data rows (0, 2, 4, ...) and **merges** with existing body styles as the lowest-priority layer. Properties set by global, column-specific, inline or conditional body styles win; the alternate row style supplies only the properties they leave unset. This means you can combine `alternateRowStyle` with `bodyStyle`:

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

### Formulas

Use Excel formulas in cells with the `formula()` function:

```kotlin
import io.clroot.excel.core.model.formula

excel {
    sheet<SummaryRow>("Summary") {
        column("Label") { it.label }
        column("Value") { formula("SUM(B2:B100)") }
        rows(summaryData)
    }
}.writeTo(output)
```

The leading `=` sign is optional - both `formula("SUM(A1:A10)")` and `formula("=SUM(A1:A10)")` work:

```kotlin
// These are equivalent
column("Total") { formula("SUM(A1:A10)") }
column("Total") { formula("=SUM(A1:A10)") }
```

Formulas can reference other sheets:

```kotlin
column("Grand Total") { formula("SUM(Sales!B2:B100)") }
```

**Note:** Formula validation is performed by Excel when opening the file. Invalid formulas will show as errors in Excel.

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

The [shared executable examples](../excel-dsl/src/test/kotlin/io/clroot/excel/examples/DocumentationTest.kt) cover DSL, Sequence, annotations and parsing in CI via `:excel-dsl:test`. JDK 21 or later is required.
