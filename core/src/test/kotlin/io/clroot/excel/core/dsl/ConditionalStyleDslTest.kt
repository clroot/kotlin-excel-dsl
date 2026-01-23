package io.clroot.excel.core.dsl

import io.clroot.excel.core.dsl.fontColor
import io.clroot.excel.core.style.Color
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe
import io.kotest.matchers.shouldNotBe

class ConditionalStyleDslTest : DescribeSpec({

    describe("conditionalStyle DSL") {
        it("column에 conditionalStyle을 정의할 수 있다") {
            data class Product(val name: String, val price: Int)

            val document =
                excel {
                    sheet<Product>("상품") {
                        column("상품명") { it.name }
                        column("가격", conditionalStyle = { value: Int? ->
                            when {
                                value == null -> null
                                value < 0 -> fontColor(Color.RED)
                                value > 1000000 -> fontColor(Color.GREEN)
                                else -> null
                            }
                        }) { it.price }
                        rows(listOf(Product("테스트", 100)))
                    }
                }

            val column = document.sheets[0].columns[1]
            column.conditionalStyle shouldNotBe null
        }

        it("conditionalStyle과 bodyStyle을 함께 사용할 수 있다") {
            data class Product(val name: String, val price: Int)

            val document =
                excel {
                    sheet<Product>("상품") {
                        column(
                            "가격",
                            bodyStyle = { bold() },
                            conditionalStyle = { value: Int? ->
                                if (value != null && value < 0) fontColor(Color.RED) else null
                            },
                        ) { it.price }
                        rows(listOf(Product("테스트", -100)))
                    }
                }

            val column = document.sheets[0].columns[0]
            column.bodyStyle?.bold shouldBe true
            column.conditionalStyle shouldNotBe null
        }
    }
})
