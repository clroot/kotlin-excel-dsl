package io.clroot.excel.render

import io.clroot.excel.core.dsl.CellStyleBuilder.Companion.fontColor
import io.clroot.excel.core.dsl.excel
import io.clroot.excel.core.style.Color
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream

class ConditionalStyleRenderTest : DescribeSpec({

    describe("conditionalStyle rendering") {
        it("applies red font color to negative values") {
            data class Transaction(val description: String, val amount: Int)

            val transactions =
                listOf(
                    Transaction("Income", 1000000),
                    Transaction("Expense", -500000),
                    Transaction("Income", 200000),
                )

            val document =
                excel {
                    sheet<Transaction>("Transactions") {
                        column("Description") { it.description }
                        column("Amount", conditionalStyle = { value: Int? ->
                            if (value != null && value < 0) fontColor(Color.RED) else null
                        }) { it.amount }
                        rows(transactions)
                    }
                }

            val output = ByteArrayOutputStream()
            document.writeTo(output)

            XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                val sheet = workbook.getSheetAt(0)

                // First row (1000000): no conditional style - default styling
                val positiveCell = sheet.getRow(1).getCell(1)
                positiveCell.numericCellValue shouldBe 1000000.0

                // Second row (-500000): red font
                val negativeCell = sheet.getRow(2).getCell(1)
                negativeCell.numericCellValue shouldBe -500000.0
                val font = workbook.getFontAt(negativeCell.cellStyle.fontIndex)
                val xssfFont = font as org.apache.poi.xssf.usermodel.XSSFFont
                xssfFont.xssfColor?.rgb?.let { rgb ->
                    rgb[0] shouldBe 0xFF.toByte() // Red
                    rgb[1] shouldBe 0x00.toByte() // Green
                    rgb[2] shouldBe 0x00.toByte() // Blue
                }
            }
        }

        it("merges bodyStyle and conditionalStyle") {
            data class Product(val name: String, val price: Int)

            val products =
                listOf(
                    Product("ProductA", 500),
                    Product("ProductB", -100),
                )

            val document =
                excel {
                    sheet<Product>("Products") {
                        column(
                            "Price",
                            bodyStyle = { bold() },
                            conditionalStyle = { value: Int? ->
                                if (value != null && value < 0) fontColor(Color.RED) else null
                            },
                        ) { it.price }
                        rows(products)
                    }
                }

            val output = ByteArrayOutputStream()
            document.writeTo(output)

            XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                val sheet = workbook.getSheetAt(0)

                // Positive value: only bodyStyle(bold) applied
                val positiveCell = sheet.getRow(1).getCell(0)
                workbook.getFontAt(positiveCell.cellStyle.fontIndex).bold shouldBe true

                // Negative value: bodyStyle(bold) + conditionalStyle(red) merged
                val negativeCell = sheet.getRow(2).getCell(0)
                val font = workbook.getFontAt(negativeCell.cellStyle.fontIndex)
                font.bold shouldBe true
                val xssfFont = font as org.apache.poi.xssf.usermodel.XSSFFont
                xssfFont.xssfColor?.rgb?.let { rgb ->
                    rgb[0] shouldBe 0xFF.toByte()
                }
            }
        }

        it("conditionalStyle has highest priority over all other styles") {
            data class Item(val value: Int)

            val items =
                listOf(
                    Item(-100),
                )

            val document =
                excel {
                    styles {
                        body {
                            fontColor(Color.BLUE)
                        }
                    }
                    sheet<Item>("Items") {
                        alternateRowStyle {
                            fontColor(Color.GREEN)
                        }
                        column(
                            "Value",
                            bodyStyle = { fontColor(Color.GRAY) },
                            conditionalStyle = { value: Int? ->
                                if (value != null && value < 0) fontColor(Color.RED) else null
                            },
                        ) { it.value }
                        rows(items)
                    }
                }

            val output = ByteArrayOutputStream()
            document.writeTo(output)

            XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                val sheet = workbook.getSheetAt(0)

                // conditionalStyle (RED) should override all others
                val cell = sheet.getRow(1).getCell(0)
                val font = workbook.getFontAt(cell.cellStyle.fontIndex) as org.apache.poi.xssf.usermodel.XSSFFont
                font.xssfColor?.rgb?.let { rgb ->
                    rgb[0] shouldBe 0xFF.toByte() // Red
                    rgb[1] shouldBe 0x00.toByte() // Green
                    rgb[2] shouldBe 0x00.toByte() // Blue
                }
            }
        }

        it("conditionalStyle returning null does not affect base style") {
            data class Item(val value: Int)

            // positive value - conditional returns null
            val items =
                listOf(
                    Item(100),
                )

            val document =
                excel {
                    sheet<Item>("Items") {
                        column(
                            "Value",
                            bodyStyle = { bold() },
                            conditionalStyle = { value: Int? ->
                                if (value != null && value < 0) fontColor(Color.RED) else null
                            },
                        ) { it.value }
                        rows(items)
                    }
                }

            val output = ByteArrayOutputStream()
            document.writeTo(output)

            XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                val sheet = workbook.getSheetAt(0)

                // Positive value: only bodyStyle(bold) applied, no red color
                val cell = sheet.getRow(1).getCell(0)
                val font = workbook.getFontAt(cell.cellStyle.fontIndex)
                font.bold shouldBe true
                // Font color should be default (black or auto)
                val xssfFont = font as org.apache.poi.xssf.usermodel.XSSFFont
                val fontColor = xssfFont.xssfColor
                // null or auto color means default black
                (fontColor == null || fontColor.isAuto) shouldBe true
            }
        }
    }
})
