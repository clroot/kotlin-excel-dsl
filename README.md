# kotlin-excel-dsl

[한국어](README.ko.md)

Kotlin DSL for creating Excel files with type safety and elegant syntax.

## Features

- **Hybrid API**: Annotations for simple cases, DSL for complex ones
- **Type Safety**: Compile-time verification of configurations
- **CSS-like Styling**: Intuitive style definitions with themes
- **Header Groups**: Multi-row headers with automatic cell merging
- **Freeze Panes / Auto Filter / Zebra Stripes**: Common Excel features
- **Streaming Support**: SXSSF for large datasets (1M+ rows)
- **Excel Parsing**: Type-safe parsing of Excel files into data classes

## Requirements

- Kotlin 2.2.0+
- JDK 21+

## Installation

```kotlin
dependencies {
    implementation("io.clroot.excel:excel-dsl:0.1.0")
}
```

## Quick Start

### DSL Approach

```kotlin
data class User(val name: String, val age: Int, val joinedAt: LocalDate)

val users = listOf(
    User("Alice", 30, LocalDate.of(2024, 1, 15)),
    User("Bob", 25, LocalDate.of(2024, 3, 20))
)

excel {
    sheet<User>("Users") {
        column("Name") { it.name }
        column("Age") { it.age }
        column("Joined") { it.joinedAt }
        rows(users)
    }
}.writeTo(FileOutputStream("users.xlsx"))
```

### Annotation Approach

```kotlin
@Excel
data class User(
    @Column("Name", order = 1) val name: String,
    @Column("Age", order = 2) val age: Int,
    @Column("Joined", order = 3) val joinedAt: LocalDate
)

excelOf(users).writeTo(FileOutputStream("users.xlsx"))
```

### With Theme

```kotlin
excel(theme = Theme.Modern) {
    sheet<User>("Users") {
        column("Name") { it.name }
        column("Age") { it.age }
        rows(users)
    }
}.writeTo(output)
```

### Parsing Excel Files

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

## Documentation

- [Usage Guide](docs/guide.md) - Styling, themes, and advanced features
- [Parsing](docs/parsing.md) - Excel file parsing configuration
- [Error Handling](docs/error-handling.md) - Exception types and handling
- [Performance](docs/performance.md) - Benchmarks and best practices

## Module Structure

```
kotlin-excel-dsl/
├── excel-dsl/   # All-in-one module (recommended)
├── core/        # Core models & DSL
├── annotation/  # @Excel, @Column processing
├── render/      # Apache POI integration
├── theme/       # Predefined themes
└── parser/      # Excel file parsing
```

## License

MIT License
