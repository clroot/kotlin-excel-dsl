package io.clroot.excel.annotation.style

import io.clroot.excel.annotation.BodyStyle
import io.clroot.excel.annotation.HeaderStyle
import io.clroot.excel.core.style.Alignment
import io.clroot.excel.core.style.Color
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.nulls.shouldBeNull
import io.kotest.matchers.nulls.shouldNotBeNull
import io.kotest.matchers.shouldBe

class StyleConverterTest : DescribeSpec({

    describe("parseHexColor") {
        context("유효한 HEX 색상") {
            it("#RRGGBB 형식을 파싱한다") {
                val color = StyleConverter.parseHexColor("#FF0000")
                color shouldBe Color(255, 0, 0)
            }

            it("RRGGBB 형식(# 없음)을 파싱한다") {
                val color = StyleConverter.parseHexColor("00FF00")
                color shouldBe Color(0, 255, 0)
            }

            it("소문자도 파싱한다") {
                val color = StyleConverter.parseHexColor("#aabbcc")
                color shouldBe Color(170, 187, 204)
            }
        }

        context("빈 문자열") {
            it("null을 반환한다") {
                StyleConverter.parseHexColor("").shouldBeNull()
            }
        }

        context("잘못된 형식") {
            it("예외를 던진다") {
                val result = runCatching { StyleConverter.parseHexColor("#GGGGGG") }
                result.isFailure shouldBe true
            }
        }
    }

    describe("resolveColor") {
        it("Hex가 있으면 Hex 우선") {
            val color = StyleConverter.resolveColor(StyleColor.BLUE, "#FF0000")
            color shouldBe Color(255, 0, 0)
        }

        it("Hex가 없으면 enum 사용") {
            val color = StyleConverter.resolveColor(StyleColor.BLUE, "")
            color shouldBe Color(0, 0, 255)
        }

        it("둘 다 없으면 null") {
            val color = StyleConverter.resolveColor(StyleColor.NONE, "")
            color.shouldBeNull()
        }
    }

    describe("toCellStyle(HeaderStyle)") {
        it("모든 기본값이면 null 반환") {
            @HeaderStyle
            class Dummy

            val annotation = Dummy::class.annotations.first { it is HeaderStyle } as HeaderStyle
            StyleConverter.toCellStyle(annotation).shouldBeNull()
        }

        it("bold가 true면 CellStyle 반환") {
            @HeaderStyle(bold = true)
            class Dummy

            val annotation = Dummy::class.annotations.first { it is HeaderStyle } as HeaderStyle
            val style = StyleConverter.toCellStyle(annotation)
            style.shouldNotBeNull()
            style.bold shouldBe true
        }

        it("backgroundColor enum이 설정되면 CellStyle에 반영") {
            @HeaderStyle(backgroundColor = StyleColor.BLUE)
            class Dummy

            val annotation = Dummy::class.annotations.first { it is HeaderStyle } as HeaderStyle
            val style = StyleConverter.toCellStyle(annotation)
            style.shouldNotBeNull()
            style.backgroundColor shouldBe Color(0, 0, 255)
        }

        it("backgroundColorHex가 설정되면 enum보다 우선") {
            @HeaderStyle(backgroundColor = StyleColor.BLUE, backgroundColorHex = "#FF0000")
            class Dummy

            val annotation = Dummy::class.annotations.first { it is HeaderStyle } as HeaderStyle
            val style = StyleConverter.toCellStyle(annotation)
            style.shouldNotBeNull()
            style.backgroundColor shouldBe Color(255, 0, 0)
        }
    }

    describe("toCellStyle(BodyStyle)") {
        it("모든 기본값이면 null 반환") {
            @BodyStyle
            class Dummy

            val annotation = Dummy::class.annotations.first { it is BodyStyle } as BodyStyle
            StyleConverter.toCellStyle(annotation).shouldBeNull()
        }

        it("alignment가 설정되면 CellStyle에 반영") {
            @BodyStyle(alignment = StyleAlignment.CENTER)
            class Dummy

            val annotation = Dummy::class.annotations.first { it is BodyStyle } as BodyStyle
            val style = StyleConverter.toCellStyle(annotation)
            style.shouldNotBeNull()
            style.alignment shouldBe Alignment.CENTER
        }
    }
})
