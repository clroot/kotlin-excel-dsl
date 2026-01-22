package io.clroot.excel.core.model

import io.clroot.excel.core.style.CellStyle
import io.clroot.excel.core.style.Color
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe
import io.kotest.matchers.shouldNotBe

class ConditionalStyleTest :
    DescribeSpec({

        describe("ConditionalStyle") {
            context("evaluate") {
                it("조건이 참이면 스타일을 반환한다") {
                    val conditionalStyle =
                        ConditionalStyle<Int> { value ->
                            if (value < 0) {
                                CellStyle(fontColor = Color.RED)
                            } else {
                                null
                            }
                        }

                    val result = conditionalStyle.evaluate(-100)

                    result shouldNotBe null
                    result?.fontColor shouldBe Color.RED
                }

                it("조건이 거짓이면 null을 반환한다") {
                    val conditionalStyle =
                        ConditionalStyle<Int> { value ->
                            if (value < 0) {
                                CellStyle(fontColor = Color.RED)
                            } else {
                                null
                            }
                        }

                    val result = conditionalStyle.evaluate(100)

                    result shouldBe null
                }

                it("여러 조건을 when으로 처리할 수 있다") {
                    val conditionalStyle =
                        ConditionalStyle<Int> { value ->
                            when {
                                value < 0 -> CellStyle(fontColor = Color.RED)
                                value > 1000000 -> CellStyle(fontColor = Color.GREEN, bold = true)
                                else -> null
                            }
                        }

                    conditionalStyle.evaluate(-50)?.fontColor shouldBe Color.RED
                    conditionalStyle.evaluate(2000000)?.fontColor shouldBe Color.GREEN
                    conditionalStyle.evaluate(2000000)?.bold shouldBe true
                    conditionalStyle.evaluate(500) shouldBe null
                }

                it("null 값을 처리할 수 있다") {
                    val conditionalStyle =
                        ConditionalStyle<String?> { value ->
                            if (value.isNullOrEmpty()) {
                                CellStyle(backgroundColor = Color.LIGHT_GRAY, italic = true)
                            } else {
                                null
                            }
                        }

                    conditionalStyle.evaluate(null)?.backgroundColor shouldBe Color.LIGHT_GRAY
                    conditionalStyle.evaluate("")?.italic shouldBe true
                    conditionalStyle.evaluate("text") shouldBe null
                }

                it("복잡한 객체 타입을 처리할 수 있다") {
                    data class Product(val name: String, val price: Int, val inStock: Boolean)

                    val conditionalStyle =
                        ConditionalStyle<Product> { product ->
                            when {
                                !product.inStock -> CellStyle(fontColor = Color.GRAY, italic = true)
                                product.price > 10000 -> CellStyle(fontColor = Color.BLUE, bold = true)
                                else -> null
                            }
                        }

                    val outOfStock = Product("Item A", 5000, false)
                    conditionalStyle.evaluate(outOfStock)?.fontColor shouldBe Color.GRAY
                    conditionalStyle.evaluate(outOfStock)?.italic shouldBe true

                    val expensive = Product("Item B", 15000, true)
                    conditionalStyle.evaluate(expensive)?.fontColor shouldBe Color.BLUE
                    conditionalStyle.evaluate(expensive)?.bold shouldBe true

                    val normal = Product("Item C", 5000, true)
                    conditionalStyle.evaluate(normal) shouldBe null
                }
            }
        }
    })
