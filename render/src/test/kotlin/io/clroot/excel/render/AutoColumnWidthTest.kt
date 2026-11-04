package io.clroot.excel.render

import io.clroot.excel.core.dsl.excel
import io.clroot.excel.core.model.ExcelDocument
import io.clroot.excel.core.model.formula
import io.clroot.excel.core.style.CellStyle
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayInputStream
import java.math.BigDecimal
import java.time.LocalDate
import java.time.LocalDateTime

/** Expected POI width for a display string of [chars] ASCII characters (with padding and minimum). */
private fun widthFor(chars: Int): Int = ColumnWidthCalculator.toPoiWidth(chars.toDouble())

private fun ExcelDocument.columnWidths(): List<Int> =
    XSSFWorkbook(ByteArrayInputStream(toByteArray())).use { workbook ->
        val sheet = workbook.getSheetAt(0)
        val columnCount = sheet.getRow(0).lastCellNum.toInt()
        (0 until columnCount).map { sheet.getColumnWidth(it) }
    }

class AutoColumnWidthTest :
    DescribeSpec({

        data class Row(val value: Any?)

        fun widthOf(
            value: Any?,
            format: String? = null,
            style: CellStyle? = null,
        ): Int =
            excel {
                sheet<Row>("S") {
                    if (style != null) {
                        column("V", format = format, conditionalStyle = { _: Any? -> style }) { it.value }
                    } else {
                        column("V", format = format) { it.value }
                    }
                    rows(listOf(Row(value)))
                }
            }.columnWidths().single()

        describe("formula cells") {
            it("use a fixed estimate instead of the Formula toString()") {
                widthOf(formula("SUM(A2:A100)")) shouldBe widthFor(ColumnWidthTracker.FORMULA_WIDTH_CHARS)
            }

            it("do not grow with long expressions") {
                val long = "IF(AND(A2>0,B2>0),SUM(A2:A100)*AVERAGE(B2:B100),VLOOKUP(C2,D:E,2,FALSE))"
                widthOf(formula(long)) shouldBe widthFor(ColumnWidthTracker.FORMULA_WIDTH_CHARS)
            }
        }

        describe("number cells without a number format") {
            it("measure large Doubles in plain notation") {
                // "10000000", not "1.0E7"
                widthOf(10_000_000.0) shouldBe widthFor(8)
                // "123456789012.5", not "1.234567890125E11"
                widthOf(123_456_789_012.5) shouldBe widthFor(14)
            }

            it("drop the trailing .0 of integral Doubles") {
                // "12345678"
                widthOf(12_345_678.0) shouldBe widthFor(8)
            }

            it("measure BigDecimal in plain notation") {
                // "1000000000000"
                widthOf(BigDecimal("1E+12")) shouldBe widthFor(13)
            }

            it("measure small Doubles as displayed") {
                widthOf(0.000001) shouldBe widthFor("0.000001".length)
            }
        }

        describe("number cells with a number format") {
            it("include thousands separators from a column format") {
                // "1,234,567,890"
                widthOf(1_234_567_890L, format = "#,##0") shouldBe widthFor(13)
            }

            it("include decimals from the applied format") {
                // "1,234,567.50"
                widthOf(1_234_567.5, format = "#,##0.00") shouldBe widthFor(12)
            }

            it("use a numberFormat from a conditional style") {
                // "12,345,678,901"
                widthOf(12_345_678_901L, style = CellStyle(numberFormat = "#,##0")) shouldBe widthFor(14)
            }
        }

        describe("date and time cells") {
            it("measure LocalDateTime with the default date-time format") {
                // "2026-11-04 09:30:15", not "2026-11-04T09:30:15.123"
                widthOf(LocalDateTime.of(2026, 11, 4, 9, 30, 15, 123_000_000)) shouldBe widthFor(19)
            }

            it("measure LocalDate with the default date format") {
                // "2026-11-04"
                widthOf(LocalDate.of(2026, 11, 4)) shouldBe widthFor(10)
            }

            it("measure dates with the applied date format") {
                // "2026/11/04"
                widthOf(LocalDateTime.of(2026, 11, 4, 9, 30, 15), format = "yyyy/mm/dd") shouldBe widthFor(10)
                // "04 November 2026 09:30"
                widthOf(LocalDateTime.of(2026, 11, 4, 9, 30), format = "dd mmmm yyyy hh:mm") shouldBe widthFor(22)
            }
        }

        describe("other cells") {
            it("measure booleans as TRUE/FALSE") {
                widthOf(true) shouldBe widthFor(4)
                widthOf(false) shouldBe widthFor(5)
            }

            it("keep measuring strings by their characters, including CJK") {
                widthOf("가나다라마바사") shouldBe widthFor(14)
            }
        }
    })
