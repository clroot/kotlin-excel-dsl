package io.clroot.excel.parser

import io.kotest.assertions.throwables.shouldThrow
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe
import io.kotest.matchers.string.shouldContain
import io.kotest.matchers.types.shouldBeInstanceOf
import java.math.BigDecimal
import java.time.LocalDate
import java.time.LocalDateTime

class CellConverterTest : DescribeSpec({

    val converter = CellConverter()

    describe("정확한 정수와 Boolean 변환") {
        it("Long 텍스트와 경계값을 정확히 보존한다") {
            listOf(Long.MIN_VALUE, Long.MAX_VALUE, 9007199254740993L).forEach {
                converter.convert(it.toString(), Long::class) shouldBe it
            }
            listOf(Int.MIN_VALUE, Int.MAX_VALUE).forEach {
                converter.convert(it.toString(), Int::class) shouldBe it
            }
            converter.convert("42.0", Int::class) shouldBe 42
            converter.convert("4.2e1", Long::class) shouldBe 42L
        }

        it("범위 초과와 소수 및 비유한 값을 거부한다") {
            listOf("2147483648", "-2147483649", Long.MAX_VALUE).forEach {
                shouldThrow<IllegalArgumentException> { converter.convert(it, Int::class) }
            }
            listOf("9223372036854775808", "-9223372036854775809", Long.MAX_VALUE.toDouble()).forEach {
                shouldThrow<IllegalArgumentException> { converter.convert(it, Long::class) }
            }
            listOf("1.9", -1.9, "NaN", "Infinity", Double.NaN, Double.POSITIVE_INFINITY).forEach {
                shouldThrow<IllegalArgumentException> { converter.convert(it, Int::class) }
                shouldThrow<IllegalArgumentException> { converter.convert(it, Long::class) }
            }
        }

        it("Boolean 숫자는 0과 1만 허용한다") {
            converter.convert(0.0, Boolean::class) shouldBe false
            converter.convert(1.0, Boolean::class) shouldBe true
            listOf("typo", "1", "0", 2, -1, 0.5, Double.NaN, Double.POSITIVE_INFINITY).forEach {
                shouldThrow<IllegalArgumentException> { converter.convert(it, Boolean::class) }
            }
        }

        it("빈 값과 사용자 지정 변환의 우선순위를 유지한다") {
            converter.convert(" ", Long::class, isNullable = true) shouldBe null
            converter.convert(null, Boolean::class, isNullable = true) shouldBe null
            shouldThrow<IllegalArgumentException> { converter.convert(" ", Int::class) }
            val custom = CellConverter(customConverters = mapOf(Int::class to { _: Any? -> 7 }))
            custom.convert(null, Int::class) shouldBe 7
            custom.convert("1.9", Int::class) shouldBe 7
        }
    }

    describe("String 변환") {
        it("문자열을 그대로 반환한다") {
            converter.convert("hello", String::class) shouldBe "hello"
        }

        it("숫자를 문자열로 변환한다") {
            converter.convert(123.0, String::class) shouldBe "123"
        }

        it("null은 예외를 발생시킨다 (nullable=false)") {
            val exception =
                shouldThrow<IllegalArgumentException> {
                    converter.convert(null, String::class, isNullable = false)
                }
            exception.message shouldContain "does not allow null"
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

            val customConverter =
                CellConverter(
                    customConverters =
                        mapOf(
                            Money::class to { value -> Money((value?.toString() ?: "0").toBigDecimal()) },
                        ),
                )

            customConverter.convert("100.50", Money::class) shouldBe Money(BigDecimal("100.50"))
        }
    }
})
