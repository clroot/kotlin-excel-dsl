package io.clroot.excel.render

import io.clroot.excel.annotation.Column
import io.clroot.excel.annotation.Excel
import io.clroot.excel.annotation.excelOf
import io.clroot.excel.core.dsl.excel
import io.clroot.excel.core.model.chars
import io.clroot.excel.core.style.Alignment
import io.clroot.excel.core.style.Color
import io.clroot.excel.theme.Theme
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import java.time.LocalDate

/**
 * Large dataset streaming tests to verify performance and memory efficiency.
 *
 * These tests ensure the library can handle 100K+ rows without running out of memory
 * by using Apache POI's SXSSF (Streaming Usermodel API).
 */
class LargeDatasetTest :
    DescribeSpec({

        describe("대용량 데이터 스트리밍 (DSL 방식)") {
            it("10만 행 파일을 생성한다") {
                val rowCount = 100_000

                val document =
                    excel {
                        sheet<LargeDataRow>("Data") {
                            column("ID", width = 10.chars) { it.id }
                            column("Name", width = 20.chars) { it.name }
                            column("Value", width = 15.chars) { it.value }
                            column("Date") { it.date }
                            rows(generateLargeData(rowCount))
                        }
                    }

                val output = ByteArrayOutputStream()
                document.writeTo(output)

                // Verify file is not empty and has reasonable size
                val bytes = output.toByteArray()
                (bytes.size > 1_000_000) shouldBe true
            }

            it("10만 행 파일의 첫 번째와 마지막 행이 올바르다") {
                val rowCount = 100_000

                val document =
                    excel {
                        sheet<LargeDataRow>("Data") {
                            column("ID") { it.id }
                            column("Name") { it.name }
                            column("Value") { it.value }
                            rows(generateLargeData(rowCount))
                        }
                    }

                val output = ByteArrayOutputStream()
                document.writeTo(output)

                // SXSSF로 생성된 파일은 XSSFWorkbook으로 읽을 수 있음
                XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                    val sheet = workbook.getSheetAt(0)

                    // 헤더 행
                    sheet.getRow(0).getCell(0).stringCellValue shouldBe "ID"

                    // 첫 번째 데이터 행
                    sheet.getRow(1).getCell(0).numericCellValue shouldBe 1.0
                    sheet.getRow(1).getCell(1).stringCellValue shouldBe "Name-1"

                    // 마지막 데이터 행
                    val lastRow = sheet.getRow(rowCount)
                    lastRow.getCell(0).numericCellValue shouldBe rowCount.toDouble()
                    lastRow.getCell(1).stringCellValue shouldBe "Name-$rowCount"
                }
            }
        }

        describe("대용량 데이터 스트리밍 (어노테이션 방식)") {
            it("10만 행을 어노테이션 방식으로 생성한다") {
                val rowCount = 100_000
                val data =
                    (1..rowCount).map { i ->
                        AnnotatedLargeDataRow(
                            id = i.toLong(),
                            name = "Name-$i",
                            value = i * 1.5,
                            date = LocalDate.of(2024, 1, 1).plusDays(i.toLong() % 365),
                        )
                    }

                val document = excelOf(data)
                val output = ByteArrayOutputStream()
                document.writeTo(output)

                XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                    val sheet = workbook.getSheetAt(0)
                    sheet.lastRowNum shouldBe rowCount
                    sheet.getRow(1).getCell(0).numericCellValue shouldBe 1.0
                    sheet.getRow(rowCount).getCell(0).numericCellValue shouldBe rowCount.toDouble()
                }
            }
        }

        describe("대용량 데이터 + 스타일") {
            it("10만 행에 테마를 적용해도 스타일 제한에 걸리지 않는다") {
                val rowCount = 100_000

                val document =
                    excel(theme = Theme.Modern) {
                        sheet<LargeDataRow>("Data") {
                            column("ID") { it.id }
                            column("Name") { it.name }
                            column("Value") { it.value }
                            rows(generateLargeData(rowCount))
                        }
                    }

                val output = ByteArrayOutputStream()
                // POI는 워크북당 최대 약 64,000개 스타일 지원
                // 테마 적용 시에도 스타일이 재사용되어야 함
                document.writeTo(output)

                XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                    val sheet = workbook.getSheetAt(0)
                    sheet.lastRowNum shouldBe rowCount // header + data rows
                    (workbook.numCellStyles < 100) shouldBe true
                    val firstStyle = sheet.getRow(1).getCell(0).cellStyle.index
                    sheet.getRow(rowCount).getCell(0).cellStyle.index shouldBe firstStyle
                }
            }

            it("컬럼별 스타일 적용 시에도 대용량 데이터를 처리한다") {
                val rowCount = 50_000 // 스타일이 많으면 조금 줄임

                val document =
                    excel {
                        styles {
                            header {
                                backgroundColor(Color.LIGHT_GRAY)
                                bold()
                            }
                            column("Value") {
                                body {
                                    align(Alignment.RIGHT)
                                }
                            }
                        }
                        sheet<LargeDataRow>("Data") {
                            column("ID") { it.id }
                            column("Name") { it.name }
                            column("Value") { it.value }
                            rows(generateLargeData(rowCount))
                        }
                    }

                val output = ByteArrayOutputStream()
                document.writeTo(output)

                XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                    val sheet = workbook.getSheetAt(0)
                    sheet.lastRowNum shouldBe rowCount
                }
            }
        }

        describe("멀티 시트 대용량") {
            it("여러 시트에 각각 5만 행씩 생성한다") {
                val rowCountPerSheet = 50_000

                val document =
                    excel {
                        sheet<LargeDataRow>("Sheet1") {
                            column("ID") { it.id }
                            column("Name") { it.name }
                            rows(generateLargeData(rowCountPerSheet))
                        }
                        sheet<LargeDataRow>("Sheet2") {
                            column("ID") { it.id }
                            column("Value") { it.value }
                            rows(generateLargeData(rowCountPerSheet))
                        }
                    }

                val output = ByteArrayOutputStream()
                document.writeTo(output)
                XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                    workbook.numberOfSheets shouldBe 2
                    for (sheet in workbook) {
                        sheet.lastRowNum shouldBe rowCountPerSheet
                        sheet.getRow(1).getCell(0).numericCellValue shouldBe 1.0
                        sheet.getRow(rowCountPerSheet).getCell(0).numericCellValue shouldBe rowCountPerSheet.toDouble()
                    }
                }
            }
        }
    })

data class LargeDataRow(
    val id: Long,
    val name: String,
    val value: Double,
    val date: LocalDate,
)

@Excel
data class AnnotatedLargeDataRow(
    @Column("ID", order = 1)
    val id: Long,
    @Column("Name", order = 2)
    val name: String,
    @Column("Value", order = 3)
    val value: Double,
    @Column("Date", order = 4)
    val date: LocalDate,
)

private fun generateLargeData(count: Int): List<LargeDataRow> =
    (1..count).map { i ->
        LargeDataRow(
            id = i.toLong(),
            name = "Name-$i",
            value = i * 1.5,
            date = LocalDate.of(2024, 1, 1).plusDays(i.toLong() % 365),
        )
    }
