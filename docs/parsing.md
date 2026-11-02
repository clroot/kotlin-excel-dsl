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

## Integer and Boolean conversion policy

Built-in Int/Long converters parse text without going through Double. Values must be mathematically integral and within the target range. Decimal and exponent notation such as `42.0` and `4.2e1` are accepted; fractional values, overflow, NaN, and Infinity fail with IllegalArgumentException. Numeric cells are checked using their stored value; precision already lost in Excel cannot be recovered. Store large integer identifiers as text.

Boolean accepts native Boolean values, case-insensitive `true`/`false` strings, and numeric 0/1 only. Strings `"0"`/`"1"`, other numbers, and unrecognized strings fail. Whitespace trimming and nullable/blank handling still follow the existing configuration. Custom converters retain priority and receive the original raw value.

This changes the previous permissive behavior: fractions are no longer truncated, overflow no longer saturates/wraps, and invalid Boolean strings no longer become false. Clean source data or explicitly register a converter if your application needs the old policy, for example:

```kotlin
parseExcel<MyRow>(input) {
    converter<Int> { value -> value.toString().toDouble().toInt() }
    converter<Boolean> { value ->
        when (value) {
            is Boolean -> value
            is Number -> value.toInt() != 0
            else -> value.toString().trim().equals("true", ignoreCase = true)
        }
    }
}
```

These opt-in converters intentionally restore lossy conversion. Conversion failures remain ParseError entries with row and column context: FAIL_FAST returns after the first failing row, while COLLECT gathers errors across rows.

The [shared executable examples](../excel-dsl/src/test/kotlin/io/clroot/excel/examples/DocumentationTest.kt) cover DSL, Sequence, annotations and parsing in CI via `:excel-dsl:test`. JDK 21 or later is required.
