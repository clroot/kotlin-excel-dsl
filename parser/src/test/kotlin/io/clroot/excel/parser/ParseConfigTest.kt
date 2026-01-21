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
