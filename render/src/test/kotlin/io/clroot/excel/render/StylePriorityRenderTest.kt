package io.clroot.excel.render

import io.clroot.excel.core.dsl.excel
import io.clroot.excel.core.model.ExcelDocument
import io.clroot.excel.core.style.Alignment
import io.clroot.excel.core.style.BorderStyle
import io.clroot.excel.core.style.CellStyle
import io.clroot.excel.core.style.Color
import io.clroot.excel.theme.Theme
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayInputStream
import java.time.LocalDate
import java.time.LocalDateTime
import org.apache.poi.ss.usermodel.BorderStyle as PoiBorderStyle

private fun ExcelDocument.styleAt(
    row: Int,
    col: Int,
    assertions: XSSFCellStyle.() -> Unit,
) {
    XSSFWorkbook(ByteArrayInputStream(toByteArray())).use { workbook ->
        val style = workbook.getSheetAt(0).getRow(row).getCell(col).cellStyle as XSSFCellStyle
        style.assertions()
    }
}

private fun XSSFCellStyle.backgroundRgb(): List<Int>? = fillForegroundXSSFColor?.rgb?.map { it.toInt() and 0xFF }

private fun XSSFCellStyle.fontRgb(): List<Int>? = font.xssfColor?.rgb?.map { it.toInt() and 0xFF }

private fun Color.rgb(): List<Int> = listOf(red, green, blue)

class StylePriorityRenderTest :
    DescribeSpec({

        data class Item(val value: Int)

        val items = listOf(Item(1), Item(2))

        describe("inline + column-specific + global body styles") {
            it("merges properties and lets the higher layer win on conflicts") {
                val document =
                    excel {
                        styles {
                            body {
                                border(BorderStyle.THIN)
                                fontColor(Color.BLUE)
                                backgroundColor(Color.LIGHT_GRAY)
                            }
                            column("Value") {
                                body {
                                    align(Alignment.RIGHT)
                                    fontColor(Color.GREEN)
                                }
                            }
                        }
                        sheet<Item>("Items") {
                            column("Value", style = {
                                bold()
                                fontColor(Color.RED)
                            }) { it.value }
                            rows(items)
                        }
                    }

                document.styleAt(1, 0) {
                    borderTop shouldBe PoiBorderStyle.THIN // global
                    backgroundRgb() shouldBe Color.LIGHT_GRAY.rgb() // global
                    alignment shouldBe HorizontalAlignment.RIGHT // column-specific
                    font.bold shouldBe true // inline
                    fontRgb() shouldBe Color.RED.rgb() // inline wins over column and global
                }
            }

            it("lets an inline style turn off bold set by a lower layer") {
                val document =
                    excel {
                        styles {
                            body { bold() }
                        }
                        sheet<Item>("Items") {
                            column("Value", style = { bold(false) }) { it.value }
                            column("Other") { it.value }
                            rows(items)
                        }
                    }

                document.styleAt(1, 0) { font.bold shouldBe false }
                document.styleAt(1, 1) { font.bold shouldBe true }
            }
        }

        describe("inline + column-specific + global header styles") {
            it("merges header styles with the same rule as body styles") {
                val document =
                    excel {
                        styles {
                            header {
                                bold()
                                backgroundColor(Color.LIGHT_GRAY)
                            }
                            column("Value") {
                                header { align(Alignment.CENTER) }
                            }
                        }
                        sheet<Item>("Items") {
                            column("Value", headerStyle = { italic() }) { it.value }
                            rows(items)
                        }
                    }

                document.styleAt(0, 0) {
                    font.bold shouldBe true
                    font.italic shouldBe true
                    backgroundRgb() shouldBe Color.LIGHT_GRAY.rgb()
                    alignment shouldBe HorizontalAlignment.CENTER
                }
            }
        }

        describe("theme + global styles") {
            it("merges styles { header } over the theme header instead of replacing it") {
                val document =
                    excel(Theme.Modern) {
                        styles {
                            header { italic() }
                        }
                        sheet<Item>("Items") {
                            column("Value") { it.value }
                            rows(items)
                        }
                    }

                document.styleAt(0, 0) {
                    font.italic shouldBe true
                    font.bold shouldBe true // from theme
                    borderTop shouldBe PoiBorderStyle.THIN // from theme
                    alignment shouldBe HorizontalAlignment.CENTER // from theme
                }
            }
        }

        describe("body + alternate row styles") {
            it("keeps body properties over alternate row properties and fills in the rest") {
                val document =
                    excel {
                        styles {
                            body { backgroundColor(Color.BLUE) }
                        }
                        sheet<Item>("Items") {
                            alternateRowStyle {
                                backgroundColor(Color.LIGHT_GRAY)
                                italic()
                            }
                            column("Value") { it.value }
                            rows(items)
                        }
                    }

                // Row 1 is the first data row (alternate row)
                document.styleAt(1, 0) {
                    backgroundRgb() shouldBe Color.BLUE.rgb()
                    font.italic shouldBe true
                }
                document.styleAt(2, 0) {
                    backgroundRgb() shouldBe Color.BLUE.rgb()
                    font.italic shouldBe false
                }
            }
        }

        describe("body + conditional styles") {
            it("lets a conditional style turn off bold from the body style") {
                val document =
                    excel {
                        sheet<Item>("Items") {
                            column(
                                "Value",
                                bodyStyle = {
                                    bold()
                                    italic()
                                },
                                conditionalStyle = { value: Int? ->
                                    if (value == 1) CellStyle(bold = false) else null
                                },
                            ) { it.value }
                            rows(items)
                        }
                    }

                document.styleAt(1, 0) {
                    font.bold shouldBe false
                    font.italic shouldBe true
                }
                document.styleAt(2, 0) {
                    font.bold shouldBe true
                    font.italic shouldBe true
                }
            }
        }

        describe("date cells + user numberFormat") {
            data class Event(val date: LocalDate, val time: LocalDateTime)

            val events = listOf(Event(LocalDate.of(2024, 5, 15), LocalDateTime.of(2024, 5, 15, 9, 30)))

            it("uses numberFormat from column-specific and inline styles") {
                val document =
                    excel {
                        styles {
                            column("Date") { body { numberFormat("dd.mm.yyyy") } }
                        }
                        sheet<Event>("Events") {
                            column("Date") { it.date }
                            column("Time", style = { numberFormat("hh:mm") }) { it.time }
                            rows(events)
                        }
                    }

                document.styleAt(1, 0) { dataFormatString shouldBe "dd.mm.yyyy" }
                document.styleAt(1, 1) { dataFormatString shouldBe "hh:mm" }
            }

            it("uses numberFormat from a conditional style") {
                val document =
                    excel {
                        sheet<Event>("Events") {
                            column("Date", conditionalStyle = { _: LocalDate? ->
                                CellStyle(numberFormat = "yyyy\"Q\"q")
                            }) { it.date }
                            rows(events)
                        }
                    }

                document.styleAt(1, 0) { dataFormatString shouldBe "yyyy\"Q\"q" }
            }

            it("falls back to the default date formats when no numberFormat is set") {
                val document =
                    excel {
                        styles {
                            body { border(BorderStyle.THIN) }
                        }
                        sheet<Event>("Events") {
                            column("Date") { it.date }
                            column("Time") { it.time }
                            rows(events)
                        }
                    }

                document.styleAt(1, 0) {
                    dataFormatString shouldBe "yyyy-mm-dd"
                    borderTop shouldBe PoiBorderStyle.THIN
                }
                document.styleAt(1, 1) { dataFormatString shouldBe "yyyy-mm-dd hh:mm:ss" }
            }
        }
    })
