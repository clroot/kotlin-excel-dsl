package io.clroot.excel.core.model

import io.kotest.assertions.throwables.shouldThrow
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe

class FormulaTest :
    DescribeSpec({

        describe("Formula") {
            it("수식 문자열을 저장한다") {
                val formula = Formula("SUM(A1:A10)")
                formula.expression shouldBe "SUM(A1:A10)"
            }

            it("선행 등호가 있으면 제거한다") {
                val formula = Formula("=SUM(A1:A10)")
                formula.normalizedExpression shouldBe "SUM(A1:A10)"
            }

            it("빈 수식은 허용하지 않는다") {
                shouldThrow<IllegalArgumentException> {
                    Formula("")
                }
            }

            it("공백만 있는 수식은 허용하지 않는다") {
                shouldThrow<IllegalArgumentException> {
                    Formula("   ")
                }
            }

            it("등호만 있는 수식은 허용하지 않는다") {
                shouldThrow<IllegalArgumentException> {
                    Formula("=")
                }
            }
        }

        describe("formula 함수") {
            it("Formula 객체를 생성한다") {
                val f = formula("AVERAGE(B2:B100)")
                f.expression shouldBe "AVERAGE(B2:B100)"
            }
        }
    })
