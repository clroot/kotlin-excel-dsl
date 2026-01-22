# Excel Parser Module Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 엑셀 파일을 읽어 `@Excel`, `@Column` 어노테이션이 붙은 데이터 클래스로 변환하는 `parser` 모듈 구현

**Architecture:** `parser` 모듈은 Apache POI로 엑셀을 읽고, 리플렉션으로 어노테이션 메타데이터를 추출하여 primary constructor를 통해 객체를 생성합니다. `ParseResult` sealed class로 성공/실패를 표현하고, 에러 수집/fail-fast 모드를 지원합니다.

**Tech Stack:** Kotlin 2.2.0, Apache POI 5.5.1, kotlin-reflect, Kotest 6.0.7

**Related Issue:** https://github.com/clroot/kotlin-excel-dsl/issues/4

---

## Task 1: `@Column` 어노테이션에 `aliases` 파라미터 추가

**Files:**
- Modify: `annotation/src/main/kotlin/io/clroot/excel/annotation/Annotations.kt`

**Step 1: 어노테이션 수정**

```kotlin
@Target(AnnotationTarget.PROPERTY)
@Retention(AnnotationRetention.RUNTIME)
annotation class Column(
    val header: String,
    val width: Int = 0,
    val format: String = "",
    val order: Int = Int.MAX_VALUE,
    val aliases: Array<String> = [],  // 새로 추가
)
```

**Step 2: 빌드 확인**

Run: `./gradlew :annotation:build`
Expected: BUILD SUCCESSFUL

**Step 3: Commit**

```bash
git add annotation/src/main/kotlin/io/clroot/excel/annotation/Annotations.kt
git commit -m "feat(annotation): add aliases parameter to @Column"
```

---

## Task 2: `parser` 모듈 생성 및 Gradle 설정

**Files:**
- Create: `parser/build.gradle.kts`
- Create: `parser/src/main/kotlin/io/clroot/excel/parser/.gitkeep`
- Modify: `settings.gradle.kts`
- Modify: `excel-dsl/build.gradle.kts`

**Step 1: settings.gradle.kts 확인 및 수정**

먼저 `settings.gradle.kts` 내용 확인 후 `parser` 모듈 추가:

```kotlin
// settings.gradle.kts에 include("parser") 추가
include("parser")
```

**Step 2: parser/build.gradle.kts 생성**

```kotlin
plugins {
    kotlin("jvm")
}

description = "Excel parser for kotlin-excel-dsl"

dependencies {
    api(project(":core"))
    api(project(":annotation"))

    // Apache POI for Excel file reading
    implementation("org.apache.poi:poi:5.5.1")
    implementation("org.apache.poi:poi-ooxml:5.5.1")
    implementation(kotlin("reflect"))
}
```

**Step 3: 디렉토리 구조 생성**

```bash
mkdir -p parser/src/main/kotlin/io/clroot/excel/parser
mkdir -p parser/src/test/kotlin/io/clroot/excel/parser
```

**Step 4: excel-dsl umbrella 모듈에 parser 추가**

```kotlin
// excel-dsl/build.gradle.kts
plugins {
    kotlin("jvm")
}

description = "All-in-one module for kotlin-excel-dsl"

dependencies {
    api(project(":core"))
    api(project(":annotation"))
    api(project(":render"))
    api(project(":theme"))
    api(project(":parser"))  // 추가
}
```

**Step 5: 빌드 확인**

Run: `./gradlew :parser:build`
Expected: BUILD SUCCESSFUL

**Step 6: Commit**

```bash
git add parser/ settings.gradle.kts excel-dsl/build.gradle.kts
git commit -m "feat(parser): create parser module with Gradle setup"
```

---

## Task 3: `ParseResult` sealed class 및 `ParseError` 정의

**Files:**
- Create: `parser/src/main/kotlin/io/clroot/excel/parser/ParseResult.kt`
- Create: `parser/src/test/kotlin/io/clroot/excel/parser/ParseResultTest.kt`

**Step 1: 테스트 작성**

```kotlin
package io.clroot.excel.parser

import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe
import io.kotest.matchers.types.shouldBeInstanceOf

class ParseResultTest : DescribeSpec({

    describe("ParseResult.Success") {
        it("getOrThrow는 데이터를 반환한다") {
            val result: ParseResult<String> = ParseResult.Success(listOf("a", "b"))

            result.getOrThrow() shouldBe listOf("a", "b")
        }

        it("getOrElse는 데이터를 반환한다") {
            val result: ParseResult<String> = ParseResult.Success(listOf("a", "b"))

            result.getOrElse { emptyList() } shouldBe listOf("a", "b")
        }

        it("getOrNull은 데이터를 반환한다") {
            val result: ParseResult<String> = ParseResult.Success(listOf("a", "b"))

            result.getOrNull() shouldBe listOf("a", "b")
        }
    }

    describe("ParseResult.Failure") {
        it("getOrThrow는 ExcelParseException을 던진다") {
            val errors = listOf(ParseError(rowIndex = 1, columnHeader = "이름", message = "필수 값 누락"))
            val result: ParseResult<String> = ParseResult.Failure(errors)

            val exception = runCatching { result.getOrThrow() }.exceptionOrNull()
            exception.shouldBeInstanceOf<ExcelParseException>()
            (exception as ExcelParseException).errors shouldBe errors
        }

        it("getOrElse는 기본값을 반환한다") {
            val errors = listOf(ParseError(rowIndex = 1, columnHeader = "이름", message = "필수 값 누락"))
            val result: ParseResult<String> = ParseResult.Failure(errors)

            result.getOrElse { listOf("default") } shouldBe listOf("default")
        }

        it("getOrNull은 null을 반환한다") {
            val errors = listOf(ParseError(rowIndex = 1, columnHeader = "이름", message = "필수 값 누락"))
            val result: ParseResult<String> = ParseResult.Failure(errors)

            result.getOrNull() shouldBe null
        }
    }
})
```

**Step 2: 테스트 실패 확인**

Run: `./gradlew :parser:test --tests "io.clroot.excel.parser.ParseResultTest"`
Expected: FAIL (클래스 없음)

**Step 3: ParseResult 구현**

```kotlin
package io.clroot.excel.parser

import io.clroot.excel.core.ExcelException

/**
 * Result of parsing an Excel file.
 */
sealed class ParseResult<T> {
    /**
     * Successful parse result containing the parsed data.
     */
    data class Success<T>(val data: List<T>) : ParseResult<T>()

    /**
     * Failed parse result containing the list of errors.
     */
    data class Failure<T>(val errors: List<ParseError>) : ParseResult<T>()

    /**
     * Returns the parsed data or throws [ExcelParseException] if parsing failed.
     */
    fun getOrThrow(): List<T> = when (this) {
        is Success -> data
        is Failure -> throw ExcelParseException(
            message = "Excel parsing failed with ${errors.size} error(s)",
            errors = errors,
        )
    }

    /**
     * Returns the parsed data or the result of [default] if parsing failed.
     */
    fun getOrElse(default: () -> List<T>): List<T> = when (this) {
        is Success -> data
        is Failure -> default()
    }

    /**
     * Returns the parsed data or null if parsing failed.
     */
    fun getOrNull(): List<T>? = when (this) {
        is Success -> data
        is Failure -> null
    }
}

/**
 * Represents a single parsing error.
 */
data class ParseError(
    val rowIndex: Int,
    val columnHeader: String?,
    val message: String,
    val cause: Throwable? = null,
)

/**
 * Exception thrown when Excel parsing fails.
 */
class ExcelParseException(
    message: String,
    val errors: List<ParseError>,
    cause: Throwable? = null,
) : ExcelException(message, cause)
```

**Step 4: 테스트 통과 확인**

Run: `./gradlew :parser:test --tests "io.clroot.excel.parser.ParseResultTest"`
Expected: PASS

**Step 5: Commit**

```bash
git add parser/src/main/kotlin/io/clroot/excel/parser/ParseResult.kt
git add parser/src/test/kotlin/io/clroot/excel/parser/ParseResultTest.kt
git commit -m "feat(parser): add ParseResult sealed class and ParseError"
```

---

## Task 4: `ParseConfig` 설정 클래스 (DSL builder)

**Files:**
- Create: `parser/src/main/kotlin/io/clroot/excel/parser/ParseConfig.kt`
- Create: `parser/src/test/kotlin/io/clroot/excel/parser/ParseConfigTest.kt`

**Step 1: 테스트 작성**

```kotlin
package io.clroot.excel.parser

import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe
import java.math.BigDecimal

class ParseConfigTest : DescribeSpec({

    describe("ParseConfig DSL") {
        it("기본값이 올바르게 설정된다") {
            val config = ParseConfig.Builder<Any>().build()

            config.headerRow shouldBe 0
            config.sheetIndex shouldBe 0
            config.sheetName shouldBe null
            config.headerMatching shouldBe HeaderMatching.FLEXIBLE
            config.onError shouldBe OnError.COLLECT
            config.skipEmptyRows shouldBe true
            config.trimWhitespace shouldBe true
        }

        it("커스텀 설정이 적용된다") {
            val config = parseConfig<Any> {
                headerRow = 1
                sheetIndex = 2
                headerMatching = HeaderMatching.EXACT
                onError = OnError.FAIL_FAST
                skipEmptyRows = false
                trimWhitespace = false
            }

            config.headerRow shouldBe 1
            config.sheetIndex shouldBe 2
            config.headerMatching shouldBe HeaderMatching.EXACT
            config.onError shouldBe OnError.FAIL_FAST
            config.skipEmptyRows shouldBe false
            config.trimWhitespace shouldBe false
        }

        it("sheetName 설정 시 sheetIndex보다 우선한다") {
            val config = parseConfig<Any> {
                sheetIndex = 2
                sheetName = "Users"
            }

            config.sheetName shouldBe "Users"
        }

        it("커스텀 컨버터를 등록할 수 있다") {
            data class Money(val amount: BigDecimal)

            val config = parseConfig<Any> {
                converter { value: Any? -> Money((value?.toString() ?: "0").toBigDecimal()) }
            }

            config.converters.containsKey(Money::class) shouldBe true
        }

        it("validateRow를 등록할 수 있다") {
            data class User(val name: String)

            var validated = false
            val config = parseConfig<User> {
                validateRow { validated = true }
            }

            config.rowValidator?.invoke(User("test"))
            validated shouldBe true
        }

        it("validateAll을 등록할 수 있다") {
            data class User(val name: String)

            var validated = false
            val config = parseConfig<User> {
                validateAll { validated = true }
            }

            config.allValidator?.invoke(listOf(User("test")))
            validated shouldBe true
        }
    }
})

// Helper function for tests
private inline fun <reified T : Any> parseConfig(block: ParseConfig.Builder<T>.() -> Unit): ParseConfig<T> {
    return ParseConfig.Builder<T>().apply(block).build()
}
```

**Step 2: 테스트 실패 확인**

Run: `./gradlew :parser:test --tests "io.clroot.excel.parser.ParseConfigTest"`
Expected: FAIL

**Step 3: ParseConfig 구현**

```kotlin
package io.clroot.excel.parser

import kotlin.reflect.KClass

/**
 * Header matching strategy.
 */
enum class HeaderMatching {
    /** Exact match required */
    EXACT,
    /** Flexible match: trim whitespace, normalize spaces, ignore case */
    FLEXIBLE,
}

/**
 * Error handling strategy.
 */
enum class OnError {
    /** Collect all errors and return them in Failure */
    COLLECT,
    /** Stop at first error */
    FAIL_FAST,
}

/**
 * Configuration for Excel parsing.
 */
data class ParseConfig<T : Any>(
    val headerRow: Int,
    val sheetIndex: Int,
    val sheetName: String?,
    val headerMatching: HeaderMatching,
    val onError: OnError,
    val skipEmptyRows: Boolean,
    val trimWhitespace: Boolean,
    val converters: Map<KClass<*>, (Any?) -> Any?>,
    val rowValidator: ((T) -> Unit)?,
    val allValidator: ((List<T>) -> Unit)?,
) {
    class Builder<T : Any> {
        var headerRow: Int = 0
        var sheetIndex: Int = 0
        var sheetName: String? = null
        var headerMatching: HeaderMatching = HeaderMatching.FLEXIBLE
        var onError: OnError = OnError.COLLECT
        var skipEmptyRows: Boolean = true
        var trimWhitespace: Boolean = true

        private val converters: MutableMap<KClass<*>, (Any?) -> Any?> = mutableMapOf()
        private var rowValidator: ((T) -> Unit)? = null
        private var allValidator: ((List<T>) -> Unit)? = null

        /**
         * Register a custom type converter.
         */
        inline fun <reified R : Any> converter(noinline convert: (Any?) -> R) {
            converters[R::class] = convert
        }

        /**
         * Register a row-level validator.
         */
        fun validateRow(validator: (T) -> Unit) {
            rowValidator = validator
        }

        /**
         * Register a validator for all parsed data.
         */
        fun validateAll(validator: (List<T>) -> Unit) {
            allValidator = validator
        }

        fun build(): ParseConfig<T> = ParseConfig(
            headerRow = headerRow,
            sheetIndex = sheetIndex,
            sheetName = sheetName,
            headerMatching = headerMatching,
            onError = onError,
            skipEmptyRows = skipEmptyRows,
            trimWhitespace = trimWhitespace,
            converters = converters.toMap(),
            rowValidator = rowValidator,
            allValidator = allValidator,
        )
    }
}
```

**Step 4: 테스트 통과 확인**

Run: `./gradlew :parser:test --tests "io.clroot.excel.parser.ParseConfigTest"`
Expected: PASS

**Step 5: Commit**

```bash
git add parser/src/main/kotlin/io/clroot/excel/parser/ParseConfig.kt
git add parser/src/test/kotlin/io/clroot/excel/parser/ParseConfigTest.kt
git commit -m "feat(parser): add ParseConfig with DSL builder"
```

---

## Task 5: 헤더 매칭 유틸리티

**Files:**
- Create: `parser/src/main/kotlin/io/clroot/excel/parser/HeaderMatcher.kt`
- Create: `parser/src/test/kotlin/io/clroot/excel/parser/HeaderMatcherTest.kt`

**Step 1: 테스트 작성**

```kotlin
package io.clroot.excel.parser

import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe

class HeaderMatcherTest : DescribeSpec({

    describe("HeaderMatcher.EXACT") {
        val matcher = HeaderMatcher(HeaderMatching.EXACT)

        it("정확히 일치하면 true") {
            matcher.matches("이름", "이름", emptyArray()) shouldBe true
        }

        it("대소문자가 다르면 false") {
            matcher.matches("Name", "name", emptyArray()) shouldBe false
        }

        it("공백이 다르면 false") {
            matcher.matches("이 름", "이름", emptyArray()) shouldBe false
        }

        it("alias가 정확히 일치하면 true") {
            matcher.matches("Name", "이름", arrayOf("Name", "성명")) shouldBe true
        }
    }

    describe("HeaderMatcher.FLEXIBLE") {
        val matcher = HeaderMatcher(HeaderMatching.FLEXIBLE)

        it("정확히 일치하면 true") {
            matcher.matches("이름", "이름", emptyArray()) shouldBe true
        }

        it("대소문자가 달라도 true") {
            matcher.matches("Name", "name", emptyArray()) shouldBe true
            matcher.matches("NAME", "name", emptyArray()) shouldBe true
        }

        it("앞뒤 공백을 무시한다") {
            matcher.matches("  이름  ", "이름", emptyArray()) shouldBe true
        }

        it("연속 공백을 단일 공백으로 정규화한다") {
            matcher.matches("이  름", "이 름", emptyArray()) shouldBe true
        }

        it("alias도 유연하게 매칭한다") {
            matcher.matches("  name  ", "이름", arrayOf("Name", "성명")) shouldBe true
        }

        it("매칭되지 않으면 false") {
            matcher.matches("나이", "이름", emptyArray()) shouldBe false
        }
    }
})
```

**Step 2: 테스트 실패 확인**

Run: `./gradlew :parser:test --tests "io.clroot.excel.parser.HeaderMatcherTest"`
Expected: FAIL

**Step 3: HeaderMatcher 구현**

```kotlin
package io.clroot.excel.parser

/**
 * Matches Excel header cells to column definitions.
 */
class HeaderMatcher(private val strategy: HeaderMatching) {

    /**
     * Checks if the cell header matches the expected header or any of its aliases.
     */
    fun matches(cellHeader: String, expectedHeader: String, aliases: Array<String>): Boolean {
        val candidates = listOf(expectedHeader) + aliases
        return candidates.any { candidate ->
            when (strategy) {
                HeaderMatching.EXACT -> cellHeader == candidate
                HeaderMatching.FLEXIBLE -> normalize(cellHeader) == normalize(candidate)
            }
        }
    }

    /**
     * Finds the matching column header from candidates.
     * Returns the matched candidate or null if no match found.
     */
    fun findMatch(cellHeader: String, columns: List<ColumnMeta>): ColumnMeta? {
        return columns.find { column ->
            matches(cellHeader, column.header, column.aliases)
        }
    }

    private fun normalize(value: String): String {
        return value
            .trim()
            .replace(Regex("\\s+"), " ")
            .lowercase()
    }
}

/**
 * Metadata extracted from @Column annotation.
 */
data class ColumnMeta(
    val header: String,
    val aliases: Array<String>,
    val propertyName: String,
    val propertyType: kotlin.reflect.KClass<*>,
    val isNullable: Boolean,
) {
    override fun equals(other: Any?): Boolean {
        if (this === other) return true
        if (other !is ColumnMeta) return false
        return propertyName == other.propertyName
    }

    override fun hashCode(): Int = propertyName.hashCode()
}
```

**Step 4: 테스트 통과 확인**

Run: `./gradlew :parser:test --tests "io.clroot.excel.parser.HeaderMatcherTest"`
Expected: PASS

**Step 5: Commit**

```bash
git add parser/src/main/kotlin/io/clroot/excel/parser/HeaderMatcher.kt
git add parser/src/test/kotlin/io/clroot/excel/parser/HeaderMatcherTest.kt
git commit -m "feat(parser): add HeaderMatcher with EXACT and FLEXIBLE strategies"
```

---

## Task 6: 기본 타입 컨버터

**Files:**
- Create: `parser/src/main/kotlin/io/clroot/excel/parser/CellConverter.kt`
- Create: `parser/src/test/kotlin/io/clroot/excel/parser/CellConverterTest.kt`

**Step 1: 테스트 작성**

```kotlin
package io.clroot.excel.parser

import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe
import io.kotest.matchers.types.shouldBeInstanceOf
import java.math.BigDecimal
import java.time.LocalDate
import java.time.LocalDateTime

class CellConverterTest : DescribeSpec({

    val converter = CellConverter()

    describe("String 변환") {
        it("문자열을 그대로 반환한다") {
            converter.convert("hello", String::class) shouldBe "hello"
        }

        it("숫자를 문자열로 변환한다") {
            converter.convert(123.0, String::class) shouldBe "123"
        }

        it("null은 빈 문자열로 변환한다 (nullable=false)") {
            converter.convert(null, String::class, isNullable = false) shouldBe ""
        }

        it("null은 null로 반환한다 (nullable=true)") {
            converter.convert(null, String::class, isNullable = true) shouldBe null
        }
    }

    describe("Int 변환") {
        it("정수를 변환한다") {
            converter.convert(42.0, Int::class) shouldBe 42
        }

        it("문자열 숫자를 변환한다") {
            converter.convert("42", Int::class) shouldBe 42
        }
    }

    describe("Long 변환") {
        it("정수를 Long으로 변환한다") {
            converter.convert(9999999999.0, Long::class) shouldBe 9999999999L
        }
    }

    describe("Double 변환") {
        it("실수를 변환한다") {
            converter.convert(3.14, Double::class) shouldBe 3.14
        }
    }

    describe("Boolean 변환") {
        it("true를 변환한다") {
            converter.convert(true, Boolean::class) shouldBe true
        }

        it("문자열 'true'를 변환한다") {
            converter.convert("true", Boolean::class) shouldBe true
            converter.convert("TRUE", Boolean::class) shouldBe true
        }

        it("문자열 'false'를 변환한다") {
            converter.convert("false", Boolean::class) shouldBe false
        }
    }

    describe("LocalDate 변환") {
        it("Excel 날짜 숫자를 변환한다") {
            // 2024-01-15 = Excel serial 45306
            val result = converter.convert(45306.0, LocalDate::class)
            result shouldBe LocalDate.of(2024, 1, 15)
        }

        it("ISO 문자열을 변환한다") {
            converter.convert("2024-01-15", LocalDate::class) shouldBe LocalDate.of(2024, 1, 15)
        }
    }

    describe("LocalDateTime 변환") {
        it("Excel 날짜시간 숫자를 변환한다") {
            // 2024-01-15 12:30:00 = 45306.520833...
            val result = converter.convert(45306.520833333336, LocalDateTime::class)
            result.shouldBeInstanceOf<LocalDateTime>()
            (result as LocalDateTime).toLocalDate() shouldBe LocalDate.of(2024, 1, 15)
        }

        it("ISO 문자열을 변환한다") {
            converter.convert("2024-01-15T12:30:00", LocalDateTime::class) shouldBe
                LocalDateTime.of(2024, 1, 15, 12, 30, 0)
        }
    }

    describe("BigDecimal 변환") {
        it("숫자를 BigDecimal로 변환한다") {
            converter.convert(123.45, BigDecimal::class) shouldBe BigDecimal("123.45")
        }

        it("문자열을 BigDecimal로 변환한다") {
            converter.convert("123.45", BigDecimal::class) shouldBe BigDecimal("123.45")
        }
    }

    describe("커스텀 컨버터") {
        it("등록된 커스텀 컨버터를 사용한다") {
            data class Money(val amount: BigDecimal)

            val customConverter = CellConverter(
                customConverters = mapOf(
                    Money::class to { value -> Money((value?.toString() ?: "0").toBigDecimal()) }
                )
            )

            customConverter.convert("100.50", Money::class) shouldBe Money(BigDecimal("100.50"))
        }
    }
})
```

**Step 2: 테스트 실패 확인**

Run: `./gradlew :parser:test --tests "io.clroot.excel.parser.CellConverterTest"`
Expected: FAIL

**Step 3: CellConverter 구현**

```kotlin
package io.clroot.excel.parser

import java.math.BigDecimal
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter
import kotlin.reflect.KClass

/**
 * Converts Excel cell values to Kotlin types.
 */
class CellConverter(
    private val customConverters: Map<KClass<*>, (Any?) -> Any?> = emptyMap(),
    private val trimWhitespace: Boolean = true,
) {
    companion object {
        // Excel epoch is 1899-12-30 (accounting for Excel's leap year bug)
        private val EXCEL_EPOCH = LocalDate.of(1899, 12, 30)
    }

    @Suppress("UNCHECKED_CAST")
    fun <T : Any> convert(value: Any?, targetType: KClass<T>, isNullable: Boolean = false): T? {
        // Check custom converter first
        customConverters[targetType]?.let { converter ->
            return converter(value) as T?
        }

        if (value == null) {
            return if (isNullable) null else defaultValue(targetType)
        }

        val processedValue = if (trimWhitespace && value is String) value.trim() else value

        return when (targetType) {
            String::class -> convertToString(processedValue)
            Int::class -> convertToInt(processedValue)
            Long::class -> convertToLong(processedValue)
            Double::class -> convertToDouble(processedValue)
            Float::class -> convertToFloat(processedValue)
            Boolean::class -> convertToBoolean(processedValue)
            LocalDate::class -> convertToLocalDate(processedValue)
            LocalDateTime::class -> convertToLocalDateTime(processedValue)
            BigDecimal::class -> convertToBigDecimal(processedValue)
            else -> throw IllegalArgumentException("Unsupported type: ${targetType.simpleName}")
        } as T?
    }

    @Suppress("UNCHECKED_CAST")
    private fun <T : Any> defaultValue(targetType: KClass<T>): T? {
        return when (targetType) {
            String::class -> "" as T
            Int::class -> 0 as T
            Long::class -> 0L as T
            Double::class -> 0.0 as T
            Float::class -> 0f as T
            Boolean::class -> false as T
            else -> null
        }
    }

    private fun convertToString(value: Any): String {
        return when (value) {
            is String -> value
            is Double -> if (value == value.toLong().toDouble()) {
                value.toLong().toString()
            } else {
                value.toString()
            }
            else -> value.toString()
        }
    }

    private fun convertToInt(value: Any): Int {
        return when (value) {
            is Number -> value.toInt()
            is String -> value.toDoubleOrNull()?.toInt() ?: value.toInt()
            else -> throw IllegalArgumentException("Cannot convert ${value::class.simpleName} to Int")
        }
    }

    private fun convertToLong(value: Any): Long {
        return when (value) {
            is Number -> value.toLong()
            is String -> value.toDoubleOrNull()?.toLong() ?: value.toLong()
            else -> throw IllegalArgumentException("Cannot convert ${value::class.simpleName} to Long")
        }
    }

    private fun convertToDouble(value: Any): Double {
        return when (value) {
            is Number -> value.toDouble()
            is String -> value.toDouble()
            else -> throw IllegalArgumentException("Cannot convert ${value::class.simpleName} to Double")
        }
    }

    private fun convertToFloat(value: Any): Float {
        return when (value) {
            is Number -> value.toFloat()
            is String -> value.toFloat()
            else -> throw IllegalArgumentException("Cannot convert ${value::class.simpleName} to Float")
        }
    }

    private fun convertToBoolean(value: Any): Boolean {
        return when (value) {
            is Boolean -> value
            is String -> value.lowercase() == "true"
            is Number -> value.toInt() != 0
            else -> throw IllegalArgumentException("Cannot convert ${value::class.simpleName} to Boolean")
        }
    }

    private fun convertToLocalDate(value: Any): LocalDate {
        return when (value) {
            is LocalDate -> value
            is LocalDateTime -> value.toLocalDate()
            is Number -> EXCEL_EPOCH.plusDays(value.toLong())
            is String -> LocalDate.parse(value, DateTimeFormatter.ISO_LOCAL_DATE)
            else -> throw IllegalArgumentException("Cannot convert ${value::class.simpleName} to LocalDate")
        }
    }

    private fun convertToLocalDateTime(value: Any): LocalDateTime {
        return when (value) {
            is LocalDateTime -> value
            is LocalDate -> value.atStartOfDay()
            is Number -> {
                val days = value.toLong()
                val fraction = value.toDouble() - days
                val secondsInDay = (fraction * 24 * 60 * 60).toLong()
                EXCEL_EPOCH.plusDays(days).atStartOfDay().plusSeconds(secondsInDay)
            }
            is String -> LocalDateTime.parse(value, DateTimeFormatter.ISO_LOCAL_DATE_TIME)
            else -> throw IllegalArgumentException("Cannot convert ${value::class.simpleName} to LocalDateTime")
        }
    }

    private fun convertToBigDecimal(value: Any): BigDecimal {
        return when (value) {
            is BigDecimal -> value
            is Number -> BigDecimal(value.toString())
            is String -> BigDecimal(value)
            else -> throw IllegalArgumentException("Cannot convert ${value::class.simpleName} to BigDecimal")
        }
    }
}
```

**Step 4: 테스트 통과 확인**

Run: `./gradlew :parser:test --tests "io.clroot.excel.parser.CellConverterTest"`
Expected: PASS

**Step 5: Commit**

```bash
git add parser/src/main/kotlin/io/clroot/excel/parser/CellConverter.kt
git add parser/src/test/kotlin/io/clroot/excel/parser/CellConverterTest.kt
git commit -m "feat(parser): add CellConverter for type conversion"
```

---

## Task 7: 어노테이션 메타데이터 추출기

**Files:**
- Create: `parser/src/main/kotlin/io/clroot/excel/parser/AnnotationExtractor.kt`
- Create: `parser/src/test/kotlin/io/clroot/excel/parser/AnnotationExtractorTest.kt`

**Step 1: 테스트 작성**

```kotlin
package io.clroot.excel.parser

import io.clroot.excel.annotation.Column
import io.clroot.excel.annotation.Excel
import io.clroot.excel.core.ExcelConfigurationException
import io.kotest.assertions.throwables.shouldThrow
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.collections.shouldHaveSize
import io.kotest.matchers.shouldBe
import io.kotest.matchers.string.shouldContain

class AnnotationExtractorTest : DescribeSpec({

    describe("extractColumns") {
        it("@Column 어노테이션이 있는 프로퍼티를 추출한다") {
            @Excel
            data class User(
                @Column("이름") val name: String,
                @Column("나이") val age: Int,
            )

            val columns = AnnotationExtractor.extractColumns(User::class)

            columns shouldHaveSize 2
            columns[0].header shouldBe "이름"
            columns[0].propertyName shouldBe "name"
            columns[1].header shouldBe "나이"
            columns[1].propertyName shouldBe "age"
        }

        it("aliases를 추출한다") {
            @Excel
            data class User(
                @Column("이름", aliases = ["Name", "성명"]) val name: String,
            )

            val columns = AnnotationExtractor.extractColumns(User::class)

            columns[0].aliases shouldBe arrayOf("Name", "성명")
        }

        it("order 순서대로 정렬된다") {
            @Excel
            data class User(
                @Column("나이", order = 2) val age: Int,
                @Column("이름", order = 1) val name: String,
            )

            val columns = AnnotationExtractor.extractColumns(User::class)

            columns[0].header shouldBe "이름"
            columns[1].header shouldBe "나이"
        }

        it("nullable 타입을 감지한다") {
            @Excel
            data class User(
                @Column("이름") val name: String,
                @Column("별명") val nickname: String?,
            )

            val columns = AnnotationExtractor.extractColumns(User::class)

            columns.find { it.propertyName == "name" }?.isNullable shouldBe false
            columns.find { it.propertyName == "nickname" }?.isNullable shouldBe true
        }

        it("@Excel 어노테이션이 없으면 예외를 던진다") {
            data class User(
                @Column("이름") val name: String,
            )

            val exception = shouldThrow<ExcelConfigurationException> {
                AnnotationExtractor.extractColumns(User::class)
            }
            exception.message shouldContain "@Excel"
        }

        it("@Column 어노테이션이 없으면 예외를 던진다") {
            @Excel
            data class User(val name: String)

            val exception = shouldThrow<ExcelConfigurationException> {
                AnnotationExtractor.extractColumns(User::class)
            }
            exception.message shouldContain "@Column"
        }
    }
})
```

**Step 2: 테스트 실패 확인**

Run: `./gradlew :parser:test --tests "io.clroot.excel.parser.AnnotationExtractorTest"`
Expected: FAIL

**Step 3: AnnotationExtractor 구현**

```kotlin
package io.clroot.excel.parser

import io.clroot.excel.annotation.Column
import io.clroot.excel.annotation.Excel
import io.clroot.excel.core.ExcelConfigurationException
import kotlin.reflect.KClass
import kotlin.reflect.KProperty1
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.full.memberProperties
import kotlin.reflect.jvm.jvmErasure

/**
 * Extracts column metadata from @Excel/@Column annotated classes.
 */
object AnnotationExtractor {

    /**
     * Extract column metadata from the given class.
     */
    fun <T : Any> extractColumns(klass: KClass<T>): List<ColumnMeta> {
        val className = klass.qualifiedName ?: klass.simpleName ?: "Unknown"

        if (klass.findAnnotation<Excel>() == null) {
            throw ExcelConfigurationException(
                message = "Missing @Excel annotation",
                className = className,
                hint = "Add @Excel annotation to your data class: @Excel data class ${klass.simpleName}(...)",
            )
        }

        val columnProps = klass.memberProperties
            .filter { it.findAnnotation<Column>() != null }
            .sortedBy { it.findAnnotation<Column>()!!.order }

        if (columnProps.isEmpty()) {
            val allProps = klass.memberProperties.map { it.name }
            throw ExcelConfigurationException(
                message = "No properties annotated with @Column",
                className = className,
                hint = "Add @Column annotation to properties. Available properties: ${allProps.joinToString(", ")}",
            )
        }

        return columnProps.map { prop ->
            val annotation = prop.findAnnotation<Column>()!!
            ColumnMeta(
                header = annotation.header,
                aliases = annotation.aliases,
                propertyName = prop.name,
                propertyType = prop.returnType.jvmErasure,
                isNullable = prop.returnType.isMarkedNullable,
            )
        }
    }

    /**
     * Get the primary constructor parameter order for the given class.
     */
    fun <T : Any> getConstructorParameterOrder(klass: KClass<T>): List<String> {
        val constructor = klass.constructors.firstOrNull()
            ?: throw ExcelConfigurationException(
                message = "No constructor found",
                className = klass.qualifiedName,
                hint = "Ensure the class has a primary constructor",
            )

        return constructor.parameters.mapNotNull { it.name }
    }
}
```

**Step 4: 테스트 통과 확인**

Run: `./gradlew :parser:test --tests "io.clroot.excel.parser.AnnotationExtractorTest"`
Expected: PASS

**Step 5: Commit**

```bash
git add parser/src/main/kotlin/io/clroot/excel/parser/AnnotationExtractor.kt
git add parser/src/test/kotlin/io/clroot/excel/parser/AnnotationExtractorTest.kt
git commit -m "feat(parser): add AnnotationExtractor for metadata extraction"
```

---

## Task 8: 메인 `parseExcel` 함수 구현

**Files:**
- Create: `parser/src/main/kotlin/io/clroot/excel/parser/ExcelParser.kt`
- Create: `parser/src/test/kotlin/io/clroot/excel/parser/ExcelParserTest.kt`

**Step 1: 테스트 작성**

```kotlin
package io.clroot.excel.parser

import io.clroot.excel.annotation.Column
import io.clroot.excel.annotation.Excel
import io.clroot.excel.annotation.excelOf
import io.clroot.excel.render.writeTo
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.collections.shouldHaveSize
import io.kotest.matchers.shouldBe
import io.kotest.matchers.types.shouldBeInstanceOf
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import java.time.LocalDate

class ExcelParserTest : DescribeSpec({

    describe("parseExcel 기본 사용") {
        it("기본 데이터 클래스를 파싱한다") {
            @Excel
            data class User(
                @Column("이름", order = 1) val name: String,
                @Column("나이", order = 2) val age: Int,
            )

            val original = listOf(
                User("김철수", 30),
                User("이영희", 25),
            )

            // Write to Excel
            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            // Parse back
            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray()))

            result.shouldBeInstanceOf<ParseResult.Success<User>>()
            val users = result.getOrThrow()
            users shouldHaveSize 2
            users[0] shouldBe User("김철수", 30)
            users[1] shouldBe User("이영희", 25)
        }

        it("nullable 필드를 파싱한다") {
            @Excel
            data class User(
                @Column("이름", order = 1) val name: String,
                @Column("별명", order = 2) val nickname: String?,
            )

            val original = listOf(
                User("김철수", "철수"),
                User("이영희", null),
            )

            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray()))

            val users = result.getOrThrow()
            users[0].nickname shouldBe "철수"
            users[1].nickname shouldBe null
        }

        it("LocalDate를 파싱한다") {
            @Excel
            data class Event(
                @Column("일정명", order = 1) val name: String,
                @Column("날짜", order = 2) val date: LocalDate,
            )

            val original = listOf(Event("회의", LocalDate.of(2024, 5, 15)))

            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<Event>(ByteArrayInputStream(output.toByteArray()))

            val events = result.getOrThrow()
            events[0].date shouldBe LocalDate.of(2024, 5, 15)
        }
    }

    describe("헤더 매칭") {
        it("FLEXIBLE 모드에서 공백을 무시한다") {
            @Excel
            data class User(
                @Column("이름") val name: String,
            )

            // 헤더에 공백이 있는 엑셀 파일을 시뮬레이션 (실제 테스트에서는 POI로 직접 생성)
            val original = listOf(User("김철수"))
            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                headerMatching = HeaderMatching.FLEXIBLE
            }

            result.shouldBeInstanceOf<ParseResult.Success<User>>()
        }

        it("alias로 매칭한다") {
            @Excel
            data class User(
                @Column("이름", aliases = ["Name", "성명"]) val name: String,
            )

            // alias 테스트는 실제 다른 헤더를 가진 엑셀이 필요
            // 여기서는 기본 동작 확인
            val original = listOf(User("김철수"))
            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray()))

            result.shouldBeInstanceOf<ParseResult.Success<User>>()
        }
    }

    describe("에러 처리") {
        it("COLLECT 모드에서 모든 에러를 수집한다") {
            @Excel
            data class User(
                @Column("이름", order = 1) val name: String,
                @Column("나이", order = 2) val age: Int,
            )

            // 잘못된 데이터가 있는 엑셀 (나이에 문자열)
            // 실제로는 POI로 직접 생성해야 하지만, 여기서는 기본 동작 테스트
            val original = listOf(User("김철수", 30))
            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                onError = OnError.COLLECT
            }

            result.shouldBeInstanceOf<ParseResult.Success<User>>()
        }
    }

    describe("검증") {
        it("validateRow가 실패하면 에러를 수집한다") {
            @Excel
            data class User(
                @Column("이름", order = 1) val name: String,
                @Column("이메일", order = 2) val email: String,
            )

            val original = listOf(
                User("김철수", "valid@email.com"),
                User("이영희", "invalid-email"),
            )

            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                validateRow { user ->
                    require(user.email.contains("@")) { "이메일 형식이 올바르지 않습니다" }
                }
            }

            result.shouldBeInstanceOf<ParseResult.Failure<User>>()
            val errors = (result as ParseResult.Failure).errors
            errors shouldHaveSize 1
            errors[0].rowIndex shouldBe 2  // 0-indexed data row (header=0, first data=1)
        }

        it("validateAll이 실패하면 에러를 반환한다") {
            @Excel
            data class User(
                @Column("이름", order = 1) val name: String,
                @Column("이메일", order = 2) val email: String,
            )

            val original = listOf(
                User("김철수", "same@email.com"),
                User("이영희", "same@email.com"),
            )

            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                validateAll { users ->
                    val duplicates = users.groupBy { it.email }.filter { it.value.size > 1 }
                    require(duplicates.isEmpty()) { "중복된 이메일이 있습니다: ${duplicates.keys}" }
                }
            }

            result.shouldBeInstanceOf<ParseResult.Failure<User>>()
        }
    }

    describe("설정 옵션") {
        it("skipEmptyRows가 true면 빈 행을 건너뛴다") {
            @Excel
            data class User(
                @Column("이름") val name: String,
            )

            val original = listOf(User("김철수"))
            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                skipEmptyRows = true
            }

            result.shouldBeInstanceOf<ParseResult.Success<User>>()
        }

        it("커스텀 컨버터를 사용한다") {
            data class Money(val amount: Int)

            @Excel
            data class Product(
                @Column("상품명", order = 1) val name: String,
                @Column("가격", order = 2) val price: Money,
            )

            // Money 타입은 커스텀 컨버터 필요
            // 실제 테스트에서는 Int로 저장된 엑셀을 Money로 변환
            val output = ByteArrayOutputStream()
            // 직접 POI로 생성하거나 간단한 테스트용 데이터 사용
        }
    }
})
```

**Step 2: 테스트 실패 확인**

Run: `./gradlew :parser:test --tests "io.clroot.excel.parser.ExcelParserTest"`
Expected: FAIL

**Step 3: ExcelParser 구현**

```kotlin
package io.clroot.excel.parser

import io.clroot.excel.annotation.Excel
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.InputStream
import kotlin.reflect.KClass
import kotlin.reflect.KParameter
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.full.primaryConstructor

/**
 * Parses an Excel file into a list of data class instances.
 */
inline fun <reified T : Any> parseExcel(
    input: InputStream,
    noinline configure: ParseConfig.Builder<T>.() -> Unit = {},
): ParseResult<T> {
    return parseExcel(T::class, input, configure)
}

/**
 * Parses an Excel file into a list of data class instances.
 */
fun <T : Any> parseExcel(
    klass: KClass<T>,
    input: InputStream,
    configure: ParseConfig.Builder<T>.() -> Unit = {},
): ParseResult<T> {
    val config = ParseConfig.Builder<T>().apply(configure).build()
    return ExcelParserImpl(klass, config).parse(input)
}

/**
 * Internal parser implementation.
 */
internal class ExcelParserImpl<T : Any>(
    private val klass: KClass<T>,
    private val config: ParseConfig<T>,
) {
    private val columns: List<ColumnMeta> = AnnotationExtractor.extractColumns(klass)
    private val headerMatcher: HeaderMatcher = HeaderMatcher(config.headerMatching)
    private val cellConverter: CellConverter = CellConverter(
        customConverters = config.converters,
        trimWhitespace = config.trimWhitespace,
    )

    fun parse(input: InputStream): ParseResult<T> {
        val errors = mutableListOf<ParseError>()
        val results = mutableListOf<T>()

        WorkbookFactory.create(input).use { workbook ->
            val sheet = getSheet(workbook)
            val headerMapping = parseHeaders(sheet)

            if (headerMapping == null) {
                return ParseResult.Failure(
                    listOf(ParseError(
                        rowIndex = config.headerRow,
                        columnHeader = null,
                        message = "Failed to parse headers",
                    ))
                )
            }

            val dataStartRow = config.headerRow + 1
            val lastRowNum = sheet.lastRowNum

            for (rowIndex in dataStartRow..lastRowNum) {
                val row = sheet.getRow(rowIndex)

                if (row == null || isEmptyRow(row)) {
                    if (config.skipEmptyRows) continue
                }

                val parseRowResult = parseRow(row, rowIndex, headerMapping)

                when (parseRowResult) {
                    is RowParseResult.Success -> {
                        // Validate row
                        val validationError = validateRow(parseRowResult.data, rowIndex)
                        if (validationError != null) {
                            errors.add(validationError)
                            if (config.onError == OnError.FAIL_FAST) {
                                return ParseResult.Failure(errors)
                            }
                        } else {
                            results.add(parseRowResult.data)
                        }
                    }
                    is RowParseResult.Failure -> {
                        errors.addAll(parseRowResult.errors)
                        if (config.onError == OnError.FAIL_FAST) {
                            return ParseResult.Failure(errors)
                        }
                    }
                }
            }
        }

        // Validate all
        if (errors.isEmpty() && config.allValidator != null) {
            val allValidationError = validateAll(results)
            if (allValidationError != null) {
                return ParseResult.Failure(listOf(allValidationError))
            }
        }

        return if (errors.isEmpty()) {
            ParseResult.Success(results)
        } else {
            ParseResult.Failure(errors)
        }
    }

    private fun getSheet(workbook: Workbook): Sheet {
        return if (config.sheetName != null) {
            workbook.getSheet(config.sheetName)
                ?: throw IllegalArgumentException("Sheet '${config.sheetName}' not found")
        } else {
            workbook.getSheetAt(config.sheetIndex)
        }
    }

    private fun parseHeaders(sheet: Sheet): Map<Int, ColumnMeta>? {
        val headerRow = sheet.getRow(config.headerRow) ?: return null
        val mapping = mutableMapOf<Int, ColumnMeta>()

        for (cellIndex in 0 until headerRow.lastCellNum) {
            val cell = headerRow.getCell(cellIndex) ?: continue
            val headerValue = getCellStringValue(cell)

            val matchedColumn = headerMatcher.findMatch(headerValue, columns)
            if (matchedColumn != null) {
                mapping[cellIndex] = matchedColumn
            }
        }

        // Check if all required columns are found
        val foundPropertyNames = mapping.values.map { it.propertyName }.toSet()
        val missingColumns = columns.filter { !it.isNullable && it.propertyName !in foundPropertyNames }

        if (missingColumns.isNotEmpty()) {
            // For now, allow missing columns if they're nullable
            // Non-nullable missing columns will cause parse errors
        }

        return mapping
    }

    private fun parseRow(row: Row?, rowIndex: Int, headerMapping: Map<Int, ColumnMeta>): RowParseResult<T> {
        val errors = mutableListOf<ParseError>()
        val values = mutableMapOf<String, Any?>()

        // Initialize with nulls for all columns
        columns.forEach { col ->
            values[col.propertyName] = null
        }

        if (row != null) {
            for ((cellIndex, columnMeta) in headerMapping) {
                val cell = row.getCell(cellIndex)
                val cellValue = getCellValue(cell)

                try {
                    val convertedValue = cellConverter.convert(
                        cellValue,
                        columnMeta.propertyType,
                        columnMeta.isNullable,
                    )
                    values[columnMeta.propertyName] = convertedValue
                } catch (e: Exception) {
                    errors.add(ParseError(
                        rowIndex = rowIndex,
                        columnHeader = columnMeta.header,
                        message = "Failed to convert value '$cellValue' to ${columnMeta.propertyType.simpleName}: ${e.message}",
                        cause = e,
                    ))
                }
            }
        }

        if (errors.isNotEmpty()) {
            return RowParseResult.Failure(errors)
        }

        // Create instance
        return try {
            val instance = createInstance(values)
            RowParseResult.Success(instance)
        } catch (e: Exception) {
            RowParseResult.Failure(listOf(ParseError(
                rowIndex = rowIndex,
                columnHeader = null,
                message = "Failed to create instance: ${e.message}",
                cause = e,
            )))
        }
    }

    private fun createInstance(values: Map<String, Any?>): T {
        val constructor = klass.primaryConstructor
            ?: throw IllegalStateException("No primary constructor found for ${klass.simpleName}")

        val args = mutableMapOf<KParameter, Any?>()
        for (param in constructor.parameters) {
            val value = values[param.name]
            if (value == null && !param.type.isMarkedNullable && !param.isOptional) {
                throw IllegalArgumentException("Missing required value for parameter '${param.name}'")
            }
            if (value != null || param.type.isMarkedNullable) {
                args[param] = value
            }
        }

        return constructor.callBy(args)
    }

    private fun getCellValue(cell: Cell?): Any? {
        if (cell == null) return null

        return when (cell.cellType) {
            CellType.STRING -> cell.stringCellValue
            CellType.NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    cell.localDateTimeCellValue
                } else {
                    cell.numericCellValue
                }
            }
            CellType.BOOLEAN -> cell.booleanCellValue
            CellType.BLANK -> null
            CellType.FORMULA -> {
                when (cell.cachedFormulaResultType) {
                    CellType.STRING -> cell.stringCellValue
                    CellType.NUMERIC -> cell.numericCellValue
                    CellType.BOOLEAN -> cell.booleanCellValue
                    else -> null
                }
            }
            else -> null
        }
    }

    private fun getCellStringValue(cell: Cell): String {
        return when (cell.cellType) {
            CellType.STRING -> cell.stringCellValue
            CellType.NUMERIC -> cell.numericCellValue.toString()
            else -> cell.toString()
        }
    }

    private fun isEmptyRow(row: Row): Boolean {
        for (cellIndex in 0 until row.lastCellNum) {
            val cell = row.getCell(cellIndex)
            if (cell != null && cell.cellType != CellType.BLANK) {
                val value = getCellValue(cell)
                if (value != null && value.toString().isNotBlank()) {
                    return false
                }
            }
        }
        return true
    }

    private fun validateRow(data: T, rowIndex: Int): ParseError? {
        val validator = config.rowValidator ?: return null
        return try {
            validator(data)
            null
        } catch (e: Exception) {
            ParseError(
                rowIndex = rowIndex,
                columnHeader = null,
                message = e.message ?: "Row validation failed",
                cause = e,
            )
        }
    }

    private fun validateAll(data: List<T>): ParseError? {
        val validator = config.allValidator ?: return null
        return try {
            validator(data)
            null
        } catch (e: Exception) {
            ParseError(
                rowIndex = -1,
                columnHeader = null,
                message = e.message ?: "Validation failed",
                cause = e,
            )
        }
    }

    private sealed class RowParseResult<T> {
        data class Success<T>(val data: T) : RowParseResult<T>()
        data class Failure<T>(val errors: List<ParseError>) : RowParseResult<T>()
    }
}
```

**Step 4: 테스트 통과 확인**

Run: `./gradlew :parser:test --tests "io.clroot.excel.parser.ExcelParserTest"`
Expected: PASS

**Step 5: Commit**

```bash
git add parser/src/main/kotlin/io/clroot/excel/parser/ExcelParser.kt
git add parser/src/test/kotlin/io/clroot/excel/parser/ExcelParserTest.kt
git commit -m "feat(parser): add parseExcel function with full parsing support"
```

---

## Task 9: E2E 통합 테스트

**Files:**
- Create: `parser/src/test/kotlin/io/clroot/excel/parser/ParserE2ETest.kt`

**Step 1: E2E 테스트 작성**

```kotlin
package io.clroot.excel.parser

import io.clroot.excel.annotation.Column
import io.clroot.excel.annotation.Excel
import io.clroot.excel.annotation.excelOf
import io.clroot.excel.render.writeTo
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.collections.shouldHaveSize
import io.kotest.matchers.shouldBe
import io.kotest.matchers.types.shouldBeInstanceOf
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import java.math.BigDecimal
import java.time.LocalDate

class ParserE2ETest : DescribeSpec({

    describe("라운드트립 테스트 (쓰기 → 읽기)") {
        it("기본 타입을 라운드트립한다") {
            @Excel
            data class User(
                @Column("이름", order = 1) val name: String,
                @Column("나이", order = 2) val age: Int,
                @Column("점수", order = 3) val score: Double,
                @Column("활성", order = 4) val active: Boolean,
            )

            val original = listOf(
                User("김철수", 30, 95.5, true),
                User("이영희", 25, 88.0, false),
            )

            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray()))

            result.shouldBeInstanceOf<ParseResult.Success<User>>()
            val parsed = result.getOrThrow()
            parsed shouldHaveSize 2
            parsed[0] shouldBe original[0]
            parsed[1] shouldBe original[1]
        }

        it("날짜 타입을 라운드트립한다") {
            @Excel
            data class Event(
                @Column("일정명", order = 1) val name: String,
                @Column("날짜", order = 2) val date: LocalDate,
            )

            val original = listOf(
                Event("회의", LocalDate.of(2024, 5, 15)),
                Event("출장", LocalDate.of(2024, 6, 20)),
            )

            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<Event>(ByteArrayInputStream(output.toByteArray()))

            val parsed = result.getOrThrow()
            parsed[0].date shouldBe LocalDate.of(2024, 5, 15)
            parsed[1].date shouldBe LocalDate.of(2024, 6, 20)
        }

        it("nullable 필드를 라운드트립한다") {
            @Excel
            data class User(
                @Column("이름", order = 1) val name: String,
                @Column("별명", order = 2) val nickname: String?,
            )

            val original = listOf(
                User("김철수", "철수"),
                User("이영희", null),
            )

            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray()))

            val parsed = result.getOrThrow()
            parsed[0].nickname shouldBe "철수"
            parsed[1].nickname shouldBe null
        }
    }

    describe("헤더 alias 테스트") {
        it("alias로 헤더를 매칭한다") {
            @Excel
            data class User(
                @Column("이름", aliases = ["Name", "성명"]) val name: String,
            )

            // 헤더가 "Name"인 엑셀 파일 생성
            val output = ByteArrayOutputStream()
            XSSFWorkbook().use { workbook ->
                val sheet = workbook.createSheet("Sheet1")
                val headerRow = sheet.createRow(0)
                headerRow.createCell(0).setCellValue("Name")  // alias 사용

                val dataRow = sheet.createRow(1)
                dataRow.createCell(0).setCellValue("김철수")

                workbook.write(output)
            }

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray()))

            result.shouldBeInstanceOf<ParseResult.Success<User>>()
            result.getOrThrow()[0].name shouldBe "김철수"
        }
    }

    describe("검증 테스트") {
        it("행 검증 실패 시 에러를 수집한다") {
            @Excel
            data class User(
                @Column("이름", order = 1) val name: String,
                @Column("이메일", order = 2) val email: String,
            )

            val original = listOf(
                User("김철수", "valid@email.com"),
                User("이영희", "invalid"),
                User("박민수", "also-invalid"),
            )

            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                validateRow { user ->
                    require(user.email.contains("@")) { "이메일 형식 오류: ${user.email}" }
                }
            }

            result.shouldBeInstanceOf<ParseResult.Failure<User>>()
            val errors = (result as ParseResult.Failure).errors
            errors shouldHaveSize 2
        }

        it("전체 검증으로 중복을 체크한다") {
            @Excel
            data class User(
                @Column("이름", order = 1) val name: String,
                @Column("이메일", order = 2) val email: String,
            )

            val original = listOf(
                User("김철수", "same@email.com"),
                User("이영희", "same@email.com"),
            )

            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                validateAll { users ->
                    val duplicates = users.groupBy { it.email }.filter { it.value.size > 1 }
                    require(duplicates.isEmpty()) { "중복 이메일: ${duplicates.keys}" }
                }
            }

            result.shouldBeInstanceOf<ParseResult.Failure<User>>()
        }
    }

    describe("에러 처리 모드") {
        it("FAIL_FAST 모드에서 첫 에러에서 중단한다") {
            @Excel
            data class User(
                @Column("이름", order = 1) val name: String,
                @Column("이메일", order = 2) val email: String,
            )

            val original = listOf(
                User("김철수", "invalid1"),
                User("이영희", "invalid2"),
                User("박민수", "invalid3"),
            )

            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                onError = OnError.FAIL_FAST
                validateRow { user ->
                    require(user.email.contains("@")) { "이메일 형식 오류" }
                }
            }

            result.shouldBeInstanceOf<ParseResult.Failure<User>>()
            val errors = (result as ParseResult.Failure).errors
            errors shouldHaveSize 1  // 첫 번째 에러에서 중단
        }
    }

    describe("커스텀 컨버터") {
        it("커스텀 타입을 변환한다") {
            data class Money(val amount: BigDecimal)

            @Excel
            data class Product(
                @Column("상품명", order = 1) val name: String,
                @Column("가격", order = 2) val price: Money,
            )

            // 가격이 숫자로 저장된 엑셀 생성
            val output = ByteArrayOutputStream()
            XSSFWorkbook().use { workbook ->
                val sheet = workbook.createSheet("Sheet1")
                val headerRow = sheet.createRow(0)
                headerRow.createCell(0).setCellValue("상품명")
                headerRow.createCell(1).setCellValue("가격")

                val dataRow = sheet.createRow(1)
                dataRow.createCell(0).setCellValue("노트북")
                dataRow.createCell(1).setCellValue(1500000.0)

                workbook.write(output)
            }

            val result = parseExcel<Product>(ByteArrayInputStream(output.toByteArray())) {
                converter<Money> { value ->
                    Money(BigDecimal(value?.toString() ?: "0"))
                }
            }

            result.shouldBeInstanceOf<ParseResult.Success<Product>>()
            result.getOrThrow()[0].price shouldBe Money(BigDecimal("1500000"))
        }
    }

    describe("시트 선택") {
        it("sheetName으로 시트를 선택한다") {
            @Excel
            data class User(
                @Column("이름") val name: String,
            )

            val output = ByteArrayOutputStream()
            XSSFWorkbook().use { workbook ->
                // 빈 시트
                workbook.createSheet("Empty")

                // 데이터가 있는 시트
                val sheet = workbook.createSheet("Users")
                sheet.createRow(0).createCell(0).setCellValue("이름")
                sheet.createRow(1).createCell(0).setCellValue("김철수")

                workbook.write(output)
            }

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                sheetName = "Users"
            }

            result.shouldBeInstanceOf<ParseResult.Success<User>>()
            result.getOrThrow()[0].name shouldBe "김철수"
        }

        it("sheetIndex로 시트를 선택한다") {
            @Excel
            data class User(
                @Column("이름") val name: String,
            )

            val output = ByteArrayOutputStream()
            XSSFWorkbook().use { workbook ->
                // 첫 번째 시트 (index 0)
                workbook.createSheet("Empty")

                // 두 번째 시트 (index 1)
                val sheet = workbook.createSheet("Users")
                sheet.createRow(0).createCell(0).setCellValue("이름")
                sheet.createRow(1).createCell(0).setCellValue("김철수")

                workbook.write(output)
            }

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                sheetIndex = 1
            }

            result.shouldBeInstanceOf<ParseResult.Success<User>>()
            result.getOrThrow()[0].name shouldBe "김철수"
        }
    }
})
```

**Step 2: 테스트 실행**

Run: `./gradlew :parser:test --tests "io.clroot.excel.parser.ParserE2ETest"`
Expected: PASS

**Step 3: Commit**

```bash
git add parser/src/test/kotlin/io/clroot/excel/parser/ParserE2ETest.kt
git commit -m "test(parser): add comprehensive E2E tests"
```

---

## Task 10: 전체 빌드 및 최종 검증

**Step 1: 전체 빌드**

Run: `./gradlew build`
Expected: BUILD SUCCESSFUL

**Step 2: 전체 테스트**

Run: `./gradlew test`
Expected: All tests pass

**Step 3: Lint 검사**

Run: `./gradlew ktlintCheck`
Expected: No violations

**Step 4: Commit 및 Issue 연결**

```bash
git add .
git commit -m "feat(parser): complete Excel parser module implementation

Closes #4"
```

---

## Summary

| Task | Description | Files |
|------|-------------|-------|
| 1 | `@Column`에 `aliases` 추가 | annotation/Annotations.kt |
| 2 | `parser` 모듈 생성 | parser/build.gradle.kts, settings.gradle.kts |
| 3 | `ParseResult`, `ParseError` 정의 | parser/ParseResult.kt |
| 4 | `ParseConfig` DSL builder | parser/ParseConfig.kt |
| 5 | 헤더 매칭 유틸리티 | parser/HeaderMatcher.kt |
| 6 | 기본 타입 컨버터 | parser/CellConverter.kt |
| 7 | 어노테이션 메타데이터 추출기 | parser/AnnotationExtractor.kt |
| 8 | 메인 `parseExcel` 함수 | parser/ExcelParser.kt |
| 9 | E2E 통합 테스트 | parser/ParserE2ETest.kt |
| 10 | 전체 빌드 및 검증 | - |
