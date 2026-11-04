package io.clroot.excel.core.style

import io.clroot.excel.core.dsl.CellStyleBuilder
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe

class CellStyleTest :
    DescribeSpec({

        describe("merge") {
            it("uses the other style's value when it is set") {
                val base = CellStyle(fontColor = Color.BLUE, alignment = Alignment.LEFT, numberFormat = "0")
                val other = CellStyle(fontColor = Color.RED, numberFormat = "#,##0")

                base.merge(other) shouldBe
                    CellStyle(fontColor = Color.RED, alignment = Alignment.LEFT, numberFormat = "#,##0")
            }

            it("keeps this style's bold and italic when the other style leaves them unset") {
                val base = CellStyle(bold = true, italic = true)

                base.merge(CellStyle(fontColor = Color.RED)).bold shouldBe true
                base.merge(CellStyle(fontColor = Color.RED)).italic shouldBe true
            }

            it("lets the other style turn bold and italic off") {
                val base = CellStyle(bold = true, italic = true)

                val merged = base.merge(CellStyle(bold = false, italic = false))

                merged.bold shouldBe false
                merged.italic shouldBe false
            }

            it("leaves bold and italic unset when neither style sets them") {
                val merged = CellStyle().merge(CellStyle())

                merged.bold shouldBe null
                merged.italic shouldBe null
            }
        }

        describe("binary compatibility") {
            it("keeps the primitive boolean constructor for previously compiled callers") {
                val constructor =
                    CellStyle::class.java.getConstructor(
                        Color::class.java,
                        Color::class.java,
                        Boolean::class.javaPrimitiveType,
                        Boolean::class.javaPrimitiveType,
                        Alignment::class.java,
                        BorderStyle::class.java,
                        String::class.java,
                    )

                constructor.newInstance(null, null, true, false, null, null, null) shouldBe
                    CellStyle(bold = true)
            }
        }

        describe("CellStyleBuilder") {
            it("leaves bold and italic unset by default") {
                val style = CellStyleBuilder().build()

                style.bold shouldBe null
                style.italic shouldBe null
            }

            it("sets bold and italic explicitly on or off") {
                CellStyleBuilder().apply { bold() }.build().bold shouldBe true
                CellStyleBuilder().apply { bold(false) }.build().bold shouldBe false
                CellStyleBuilder().apply { italic() }.build().italic shouldBe true
                CellStyleBuilder().apply { italic(false) }.build().italic shouldBe false
            }
        }
    })
