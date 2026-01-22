package io.clroot.excel.render

import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe

class ColumnWidthCalculatorTest :
    DescribeSpec({

        describe("isCjkCharacter") {
            it("한글을 CJK 문자로 인식한다") {
                ColumnWidthCalculator.isCjkCharacter('가') shouldBe true
                ColumnWidthCalculator.isCjkCharacter('힣') shouldBe true
                ColumnWidthCalculator.isCjkCharacter('ㄱ') shouldBe true
            }

            it("일본어를 CJK 문자로 인식한다") {
                // Hiragana
                ColumnWidthCalculator.isCjkCharacter('あ') shouldBe true
                // Katakana
                ColumnWidthCalculator.isCjkCharacter('ア') shouldBe true
                // Kanji
                ColumnWidthCalculator.isCjkCharacter('漢') shouldBe true
            }

            it("중국어를 CJK 문자로 인식한다") {
                ColumnWidthCalculator.isCjkCharacter('中') shouldBe true
                ColumnWidthCalculator.isCjkCharacter('国') shouldBe true
            }

            it("ASCII 문자를 CJK 문자로 인식하지 않는다") {
                ColumnWidthCalculator.isCjkCharacter('A') shouldBe false
                ColumnWidthCalculator.isCjkCharacter('z') shouldBe false
                ColumnWidthCalculator.isCjkCharacter('1') shouldBe false
                ColumnWidthCalculator.isCjkCharacter(' ') shouldBe false
            }
        }

        describe("calculateTextWidth") {
            it("ASCII 문자는 1의 너비를 갖는다") {
                ColumnWidthCalculator.calculateTextWidth("ABC") shouldBe 3.0
                ColumnWidthCalculator.calculateTextWidth("Hello") shouldBe 5.0
            }

            it("CJK 문자는 2의 너비를 갖는다") {
                ColumnWidthCalculator.calculateTextWidth("가나다") shouldBe 6.0
                ColumnWidthCalculator.calculateTextWidth("한글") shouldBe 4.0
            }

            it("혼합 문자열의 너비를 올바르게 계산한다") {
                // "이름" (2 CJK chars) + " - " (3 ASCII) + "Kim" (3 ASCII) = 4 + 3 + 3 = 10
                ColumnWidthCalculator.calculateTextWidth("이름 - Kim") shouldBe 10.0

                // "Hello" (5 ASCII) + "세계" (2 CJK) = 5 + 4 = 9
                ColumnWidthCalculator.calculateTextWidth("Hello세계") shouldBe 9.0
            }

            it("빈 문자열은 0의 너비를 갖는다") {
                ColumnWidthCalculator.calculateTextWidth("") shouldBe 0.0
            }
        }

        describe("calculateWidth") {
            it("컬럼 값들 중 최대 너비를 기준으로 계산한다") {
                val values = listOf("이름", "김철수", "이영희입니다")
                val width = ColumnWidthCalculator.calculateWidth(values)

                // "이영희입니다" = 6 CJK chars = 12 width, + padding 2 = 14
                // POI units = 14 * 256 = 3584
                width shouldBe (14 * 256)
            }

            it("최소 너비를 보장한다") {
                val values = listOf("A", "B")
                val width = ColumnWidthCalculator.calculateWidth(values)

                // Minimum is 8 chars
                width shouldBe (8 * 256)
            }

            it("최대 너비를 초과하지 않는다") {
                val longValue = "가".repeat(200) // 200 CJK chars = 400 width
                val values = listOf(longValue)
                val width = ColumnWidthCalculator.calculateWidth(values)

                // Maximum is 100 chars
                width shouldBe (100 * 256)
            }

            it("빈 리스트는 최소 너비를 반환한다") {
                val width = ColumnWidthCalculator.calculateWidth(emptyList())
                width shouldBe (8 * 256)
            }

            it("null 값을 처리한다") {
                val values = listOf<Any?>("이름", null, "값")
                val width = ColumnWidthCalculator.calculateWidth(values)

                // "이름" = 2 CJK = 4 width + 2 padding = 6, but minimum is 8
                // So should be at least 8 * 256
                width shouldBe (8 * 256)
            }
        }

        describe("실제 사용 사례") {
            it("한글 헤더와 영문 데이터가 혼합된 경우") {
                val values = listOf("사용자명", "Alice", "Bob", "Charlie")
                val width = ColumnWidthCalculator.calculateWidth(values)

                // "사용자명" = 4 CJK = 8 width + 2 padding = 10
                width shouldBe (10 * 256)
            }

            it("영문 헤더와 한글 데이터가 혼합된 경우") {
                val values = listOf("Name", "김철수", "이영희", "박민수")
                val width = ColumnWidthCalculator.calculateWidth(values)

                // "김철수" = 3 CJK = 6 width + 2 padding = 8
                width shouldBe (8 * 256)
            }

            it("숫자 데이터") {
                val values = listOf("금액", 1000000, 2500000, 999)
                val width = ColumnWidthCalculator.calculateWidth(values)

                // "2500000" = 7 chars + 2 padding = 9
                width shouldBe (9 * 256)
            }
        }
    })
