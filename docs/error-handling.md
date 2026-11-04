# Error Handling

[한국어](error-handling.ko.md)

The library provides detailed, context-rich error messages with helpful hints.

## Error Messages

```kotlin
// Missing @Excel annotation
excelOf(listOf(NonAnnotatedClass()))
// ExcelConfigurationException: Missing @Excel annotation [class=NonAnnotatedClass]
// Hint: Add @Excel annotation to your data class: @Excel data class NonAnnotatedClass(...)

// Missing @Column annotations
excelOf(listOf(NoColumnClass()))
// ExcelConfigurationException: No properties annotated with @Column [class=NoColumnClass]
// Hint: Add @Column annotation to properties. Available properties: name, age
```

## Exception Types

The library throws these `ExcelException` subtypes:

| Exception | Thrown when |
|-----------|-------------|
| `ExcelException` | Base type; catch it to handle every library error |
| `ExcelConfigurationException` | Annotation/configuration problems: missing `@Excel` or `@Column`, a `ConditionalStyler` whose generic type does not match the property type. Thrown by `excelOf`/`parseExcel`, or during `writeTo` for a styler mismatch |
| `ExcelWriteException` | Rendering or writing fails, e.g. a value extractor lambda throws or the output stream fails. When a data cell fails, `sheetName`, `rowIndex` and `columnIndex` are set and the cause is preserved |
| `ExcelParseException` | `ParseResult.getOrThrow()` on a failed parse result |

`ExcelConfigurationException` raised while rendering is propagated as-is; only non-library exceptions are wrapped in `ExcelWriteException`.

`ExcelDataException`, `ColumnNotFoundException` and `StyleException` are deprecated: the library never throws them. Missing columns and conversion failures during parsing are reported as `ParseResult.Failure` (see [Parsing](parsing.md)).

## Exception Hierarchy

```
ExcelException (base)
├── ExcelConfigurationException
├── ExcelWriteException
└── ExcelParseException
```

All library exceptions extend `ExcelException`, so a single catch block handles them:

```kotlin
try {
    excel {
        // ...
    }.writeTo(output)
} catch (e: ExcelException) {
    logger.error("Excel operation failed: ${e.message}")
}
```

Or handle specific exceptions:

```kotlin
try {
    excelOf(data).writeTo(output)
} catch (e: ExcelConfigurationException) {
    // Handle configuration errors (missing annotations, styler type mismatch, etc.)
} catch (e: ExcelWriteException) {
    // Handle rendering/writing errors
    logger.error("Failed at sheet=${e.sheetName}, row=${e.rowIndex}, column=${e.columnIndex}", e)
}
```

`rowIndex` and `columnIndex` are 0-based sheet indexes (header rows included); the message shows them 1-based, as in Excel:

```
ExcelWriteException: Failed to render cell: boom [sheet='Users', row=4, column=2]
```

## Argument Validation (`IllegalArgumentException`)

Invalid arguments passed to the DSL or annotations are rejected immediately with `IllegalArgumentException`, which is **not** an `ExcelException`. `catch (e: ExcelException)` does not catch these:

- `freezePane(row, col)` with a negative row or column
- `formula("")` / `Formula` with a blank expression
- An invalid `#RRGGBB` value in `backgroundColorHex`/`fontColorHex` (thrown by `excelOf`)
- `parseExcel` with a `sheetName` that does not exist

These are programming errors; fix the call site rather than catching them.
