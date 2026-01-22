# Parsing Excel Files

[한국어](parsing.ko.md)

Parse Excel files into type-safe data classes using the same annotations.

## Table of Contents

- [Basic Parsing](#basic-parsing)
- [Header Aliases](#header-aliases)
- [Parse Configuration](#parse-configuration)
- [Configuration Options](#configuration-options)
- [Supported Types](#supported-types)

## Basic Parsing

```kotlin
@Excel
data class User(
    @Column("Name") val name: String,
    @Column("Email") val email: String,
)

val result = parseExcel<User>(FileInputStream("users.xlsx"))

when (result) {
    is ParseResult.Success -> result.data.forEach { println(it) }
    is ParseResult.Failure -> result.errors.forEach { println(it.message) }
}
```

## Header Aliases

Support alternative header names for flexibility:

```kotlin
@Excel
data class User(
    @Column("Name", aliases = ["Full Name", "이름"]) val name: String,
    @Column("Email", aliases = ["E-mail", "이메일"]) val email: String,
)
```

This allows parsing Excel files with different header naming conventions.

## Parse Configuration

```kotlin
val result = parseExcel<User>(inputStream) {
    // Sheet selection
    sheetName = "Users"        // or sheetIndex = 0
    headerRow = 1              // 0-indexed
    
    // Header matching
    headerMatching = HeaderMatching.FLEXIBLE  // case-insensitive, whitespace normalized
    
    // Error handling
    onError = OnError.COLLECT  // collect all errors, or FAIL_FAST
    
    // Data processing
    skipEmptyRows = true
    trimWhitespace = true
    treatBlankAsNull = true
    
    // Validation
    validateRow { user -> 
        require(user.email.contains("@")) { "Invalid email" }
    }
    validateAll { users ->
        val duplicates = users.groupBy { it.email }.filter { it.value.size > 1 }
        require(duplicates.isEmpty()) { "Duplicate emails: ${duplicates.keys}" }
    }
    
    // Custom type converter
    converter<Money> { value -> Money(BigDecimal(value.toString())) }
}
```

## Configuration Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `sheetIndex` | `Int` | `0` | Sheet index (0-based) |
| `sheetName` | `String?` | `null` | Sheet name (overrides sheetIndex) |
| `headerRow` | `Int` | `0` | Header row index |
| `headerMatching` | `HeaderMatching` | `FLEXIBLE` | `EXACT` or `FLEXIBLE` |
| `onError` | `OnError` | `COLLECT` | `COLLECT` or `FAIL_FAST` |
| `skipEmptyRows` | `Boolean` | `true` | Skip empty rows |
| `trimWhitespace` | `Boolean` | `true` | Trim string values |
| `treatBlankAsNull` | `Boolean` | `true` | Treat blank strings as null |

### Header Matching Modes

- **EXACT**: Headers must match exactly (case-sensitive)
- **FLEXIBLE**: Case-insensitive, whitespace normalized (recommended)

### Error Handling Modes

- **COLLECT**: Collect all errors and return them in `ParseResult.Failure`
- **FAIL_FAST**: Stop at the first error

## Supported Types

The following types are supported out of the box:

- `String`
- `Int`, `Long`
- `Double`, `Float`
- `Boolean`
- `BigDecimal`
- `LocalDate`
- `LocalDateTime`

For custom types, use the `converter` function:

```kotlin
data class Money(val amount: BigDecimal)

val result = parseExcel<Product>(inputStream) {
    converter<Money> { value -> Money(BigDecimal(value.toString())) }
}
```
