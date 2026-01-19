# kotlin-excel-dsl

[한국어](README.ko.md)

Kotlin DSL for creating Excel files with type safety and elegant syntax.

## Requirements

- Kotlin 2.2.0+
- JDK 21+

## Features

- **Hybrid API**: Annotations for simple cases, DSL for complex ones
- **Type Safety**: Compile-time verification of configurations
- **CSS-like Styling**: Intuitive style definitions with column-level granularity
- **Predefined Themes**: Modern, Minimal, Classic
- **Header Groups**: Multi-row headers with automatic cell merging
- **Auto Date Formatting**: LocalDate/LocalDateTime automatic detection
- **Streaming Support**: SXSSF for large datasets
- **Detailed Error Messages**: Context-rich exceptions with helpful hints

## Quick Start

### Gradle

```kotlin
dependencies {
    // All-in-one (recommended)
    implementation("io.clroot.excel:excel-dsl:0.1.0")
}
```

Or individual modules:

```kotlin
dependencies {
    implementation("io.clroot.excel:core:0.1.0")
    implementation("io.clroot.excel:annotation:0.1.0")
    implementation("io.clroot.excel:render:0.1.0")
    implementation("io.clroot.excel:theme:0.1.0")
}
```

## Usage

```kotlin
import io.clroot.excel.*
```

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
    @Column("Name", order = 1)
    val name: String,

    @Column("Age", order = 2)
    val age: Int,

    @Column("Joined", order = 3)
    val joinedAt: LocalDate
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

### Custom Styles

#### Global Styles

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

#### Column-Level Styles (via styles DSL)

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

#### Inline Column Styles

```kotlin
excel {
    sheet<Product>("Products") {
        column("Name") { it.name }
        column("Price", style = { align(Alignment.RIGHT); bold() }) { it.price }
        rows(products)
    }
}.writeTo(output)
```

#### Separate Header and Body Styles per Column

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

**Style Priority**: Inline > Column-specific (styles DSL) > Global

### Header Groups (Multi-row Headers)

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

```kotlin
sheet<User>("Users") {
    column("Name", width = 20.chars) { it.name }
    column("Description", width = 50.chars) { it.description }
    rows(users)
}
```

### Multi-Sheet

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

## Error Handling

The library provides detailed, context-rich error messages:

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

### Exception Types

| Exception                     | Description                                  |
| ----------------------------- | -------------------------------------------- |
| `ExcelException`              | Base exception for all Excel-related errors  |
| `ExcelConfigurationException` | Configuration errors (annotations, settings) |
| `ExcelDataException`          | Data processing errors                       |
| `ExcelWriteException`         | File writing errors                          |
| `ColumnNotFoundException`     | Column lookup failures                       |
| `StyleException`              | Style application errors                     |

## Available Themes

| Theme           | Description                           |
| --------------- | ------------------------------------- |
| `Theme.Modern`  | Blue header with white text, centered |
| `Theme.Minimal` | Bold header, no background            |
| `Theme.Classic` | Gray header, medium borders           |

## Module Structure

```
kotlin-excel-dsl/
├── excel-dsl/   # All-in-one module (recommended)
├── core/        # Core models & DSL (pure Kotlin)
├── annotation/  # @Excel, @Column processing
├── render/      # Apache POI integration
└── theme/       # Predefined themes
```

## License

MIT License
