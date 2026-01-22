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

| Exception | Description |
|-----------|-------------|
| `ExcelException` | Base exception for all Excel-related errors |
| `ExcelConfigurationException` | Configuration errors (annotations, settings) |
| `ExcelDataException` | Data processing errors |
| `ExcelWriteException` | File writing errors |
| `ExcelParseException` | Excel parsing errors |
| `ColumnNotFoundException` | Column lookup failures |
| `StyleException` | Style application errors |

## Exception Hierarchy

```
ExcelException (base)
├── ExcelConfigurationException
├── ExcelDataException
├── ExcelWriteException
├── ExcelParseException
├── ColumnNotFoundException
└── StyleException
```

All exceptions extend `ExcelException`, allowing you to catch all library errors with a single catch block if needed:

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
    // Handle configuration errors (missing annotations, etc.)
} catch (e: ExcelWriteException) {
    // Handle file writing errors
}
```
