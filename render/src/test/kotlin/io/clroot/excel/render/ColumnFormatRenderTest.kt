package io.clroot.excel.render

import io.clroot.excel.annotation.Column
import io.clroot.excel.annotation.Excel
import io.clroot.excel.annotation.excelOf
import io.clroot.excel.core.dsl.excel
import io.clroot.excel.core.model.ExcelDocument
import io.clroot.excel.core.style.CellStyle
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayInputStream
import java.time.LocalDate
import java.time.LocalDateTime

@Excel
data class FormattedProduct(
    @Column("Name", order = 1)
    val name: String,
    @Column("Price", format = "#,##0", order = 2)
    val price: Int,
    @Column("Released", format = "yyyy/mm/dd", order = 3)
    val released: LocalDate,
)

private fun ExcelDocument.formatAt(
    row: Int,
    col: Int,
): String =
    XSSFWorkbook(ByteArrayInputStream(toByteArray())).use { workbook ->
        workbook.getSheetAt(0).getRow(row).getCell(col).cellStyle.dataFormatString
    }

class ColumnFormatRenderTest :
    DescribeSpec({

        data class Item(val amount: Int, val date: LocalDate, val time: LocalDateTime)

        val items = listOf(Item(1234567, LocalDate.of(2024, 5, 15), LocalDateTime.of(2024, 5, 15, 9, 30)))

        describe("DSL column(format = ...)") {
            it("applies format as the body cell number format") {
                val document =
                    excel {
                        sheet<Item>("Items") {
                            column("Amount", format = "#,##0") { it.amount }
                            rows(items)
                        }
                    }

                document.formatAt(1, 0) shouldBe "#,##0"
            }

            it("does not apply format to the header cell") {
                val document =
                    excel {
                        sheet<Item>("Items") {
                            column("Amount", format = "#,##0") { it.amount }
                            rows(items)
                        }
                    }

                document.formatAt(0, 0) shouldBe "General"
            }

            it("is applied together with inline style properties") {
                val document =
                    excel {
                        sheet<Item>("Items") {
                            column("Amount", format = "#,##0", style = { bold() }) { it.amount }
                            rows(items)
                        }
                    }

                XSSFWorkbook(ByteArrayInputStream(document.toByteArray())).use { workbook ->
                    val style = workbook.getSheetAt(0).getRow(1).getCell(0).cellStyle
                    style.dataFormatString shouldBe "#,##0"
                    style.font.bold shouldBe true
                }
            }

            it("takes precedence over numberFormat from inline, column and global body styles") {
                val document =
                    excel {
                        styles {
                            body { numberFormat("0.0") }
                            column("Amount") { body { numberFormat("0.00") } }
                        }
                        sheet<Item>("Items") {
                            column("Amount", format = "#,##0", style = { numberFormat("0.000") }) { it.amount }
                            rows(items)
                        }
                    }

                document.formatAt(1, 0) shouldBe "#,##0"
            }

            it("is overridden by a conditional style numberFormat") {
                val document =
                    excel {
                        sheet<Item>("Items") {
                            column(
                                "Amount",
                                format = "#,##0",
                                conditionalStyle = { value: Int? ->
                                    if (value != null && value > 1000) CellStyle(numberFormat = "0.00E+00") else null
                                },
                            ) { it.amount }
                            rows(items)
                        }
                    }

                document.formatAt(1, 0) shouldBe "0.00E+00"
            }

            it("uses the column format for LocalDate and LocalDateTime cells") {
                val document =
                    excel {
                        sheet<Item>("Items") {
                            column("Date", format = "yyyy/mm/dd") { it.date }
                            column("Time", format = "yyyy/mm/dd hh:mm") { it.time }
                            rows(items)
                        }
                    }

                document.formatAt(1, 0) shouldBe "yyyy/mm/dd"
                document.formatAt(1, 1) shouldBe "yyyy/mm/dd hh:mm"
            }
        }

        describe("@Column(format = ...)") {
            it("applies format to number and date body cells") {
                val document =
                    excelOf(listOf(FormattedProduct("A", 10000, LocalDate.of(2024, 1, 2))))

                document.formatAt(1, 1) shouldBe "#,##0"
                document.formatAt(1, 2) shouldBe "yyyy/mm/dd"
            }
        }
    })
