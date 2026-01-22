package io.clroot.excel.render

import io.clroot.excel.core.dsl.excel
import io.clroot.excel.core.model.formula
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream

class FormulaRenderTest : DescribeSpec({

    describe("수식 렌더링") {
        it("Formula 값을 수식으로 렌더링한다") {
            data class Row(val value: Int)

            val rows = listOf(Row(100), Row(200), Row(300))

            val document =
                excel {
                    sheet<Row>("데이터") {
                        column("값") { it.value }
                        rows(rows)
                    }
                    // 별도 시트에 합계 수식
                    sheet<Unit>("요약") {
                        column("합계") { formula("SUM(데이터!A2:A4)") }
                        rows(listOf(Unit))
                    }
                }

            val output = ByteArrayOutputStream()
            document.writeTo(output)

            XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                val summarySheet = workbook.getSheet("요약")
                val formulaCell = summarySheet.getRow(1).getCell(0)

                formulaCell.cellType shouldBe CellType.FORMULA
                formulaCell.cellFormula shouldBe "SUM(데이터!A2:A4)"
            }
        }

        it("간단한 수식을 렌더링한다") {
            data class Item(val a: Int, val b: Int, val sum: Any)

            val items =
                listOf(
                    Item(10, 20, formula("A2+B2")),
                    Item(30, 40, formula("A3+B3")),
                )

            val document =
                excel {
                    sheet<Item>("계산") {
                        column("A") { it.a }
                        column("B") { it.b }
                        column("합계") { it.sum }
                        rows(items)
                    }
                }

            val output = ByteArrayOutputStream()
            document.writeTo(output)

            XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                val sheet = workbook.getSheetAt(0)

                val cell1 = sheet.getRow(1).getCell(2)
                cell1.cellType shouldBe CellType.FORMULA
                cell1.cellFormula shouldBe "A2+B2"

                val cell2 = sheet.getRow(2).getCell(2)
                cell2.cellType shouldBe CellType.FORMULA
                cell2.cellFormula shouldBe "A3+B3"
            }
        }

        it("선행 등호가 있는 수식도 정상 처리한다") {
            data class Row(val formula: Any)

            val rows = listOf(Row(formula("=TODAY()")))

            val document =
                excel {
                    sheet<Row>("날짜") {
                        column("오늘") { it.formula }
                        rows(rows)
                    }
                }

            val output = ByteArrayOutputStream()
            document.writeTo(output)

            XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                val sheet = workbook.getSheetAt(0)
                val cell = sheet.getRow(1).getCell(0)

                cell.cellType shouldBe CellType.FORMULA
                cell.cellFormula shouldBe "TODAY()"
            }
        }
    }
})
