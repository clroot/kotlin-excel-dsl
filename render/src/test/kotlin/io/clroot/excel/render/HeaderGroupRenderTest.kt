package io.clroot.excel.render

import io.clroot.excel.core.dsl.excel
import io.clroot.excel.core.model.ExcelDocument
import io.clroot.excel.core.style.Color
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.collections.shouldContainExactlyInAnyOrder
import io.kotest.matchers.shouldBe
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream

class HeaderGroupRenderTest :
    DescribeSpec({

        data class Student(
            val no: Int,
            val name: String,
            val korean: Int,
            val english: Int,
            val note: String,
        )

        val students = listOf(Student(1, "김철수", 90, 85, "우수"))

        fun ExcelDocument.readFirstSheet(block: (XSSFSheet) -> Unit) {
            val output = ByteArrayOutputStream()
            writeTo(output)
            XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { block(it.getSheetAt(0)) }
        }

        fun XSSFSheet.text(
            row: Int,
            col: Int,
        ): String? = getRow(row)?.getCell(col)?.stringCellValue

        fun XSSFSheet.mergedRefs(): List<String> = mergedRegions.map { it.formatAsString() }

        describe("그룹 없는 컬럼과 헤더 그룹을 섞는 경우") {
            it("그룹 앞의 일반 컬럼은 두 헤더 행을 세로 병합하고 그룹 제목은 실제 컬럼 위에 놓인다") {
                excel {
                    sheet<Student>("성적표") {
                        autoFilter()
                        freezePane(row = 2, col = 1)
                        column("번호") { it.no }
                        headerGroup("성적") {
                            column("국어") { it.korean }
                            column("영어") { it.english }
                        }
                        rows(students)
                    }
                }.readFirstSheet { sheet ->
                    sheet.mergedRefs() shouldContainExactlyInAnyOrder listOf("A1:A2", "B1:C1")

                    sheet.text(0, 0) shouldBe "번호"
                    sheet.text(0, 1) shouldBe "성적"

                    sheet.text(1, 0) shouldBe "번호"
                    sheet.text(1, 1) shouldBe "국어"
                    sheet.text(1, 2) shouldBe "영어"

                    sheet.getRow(2).getCell(0).numericCellValue shouldBe 1.0
                    sheet.getRow(2).getCell(1).numericCellValue shouldBe 90.0

                    sheet.ctWorksheet.autoFilter.ref shouldBe "A2:C2"
                    sheet.paneInformation.horizontalSplitTopRow.toInt() shouldBe 2
                    sheet.paneInformation.verticalSplitLeftColumn.toInt() shouldBe 1
                }
            }

            it("그룹 사이의 일반 컬럼은 세로 병합되고 뒤 그룹의 제목 위치가 밀리지 않는다") {
                excel {
                    sheet<Student>("성적표") {
                        autoFilter()
                        freezePane(row = 2)
                        headerGroup("학생") {
                            column("번호") { it.no }
                            column("이름") { it.name }
                        }
                        column("비고") { it.note }
                        headerGroup("성적") {
                            column("국어") { it.korean }
                            column("영어") { it.english }
                        }
                        rows(students)
                    }
                }.readFirstSheet { sheet ->
                    sheet.mergedRefs() shouldContainExactlyInAnyOrder listOf("A1:B1", "C1:C2", "D1:E1")

                    sheet.text(0, 0) shouldBe "학생"
                    sheet.text(0, 2) shouldBe "비고"
                    sheet.text(0, 3) shouldBe "성적"

                    (0..4).map { sheet.text(1, it) } shouldBe listOf("번호", "이름", "비고", "국어", "영어")

                    sheet.getRow(2).getCell(2).stringCellValue shouldBe "우수"
                    sheet.getRow(2).getCell(3).numericCellValue shouldBe 90.0

                    sheet.ctWorksheet.autoFilter.ref shouldBe "A2:E2"
                    sheet.paneInformation.horizontalSplitTopRow.toInt() shouldBe 2
                }
            }

            it("그룹 뒤의 일반 컬럼은 세로 병합된다") {
                excel {
                    sheet<Student>("성적표") {
                        autoFilter()
                        freezePane(row = 2)
                        headerGroup("성적") {
                            column("국어") { it.korean }
                            column("영어") { it.english }
                        }
                        column("비고") { it.note }
                        rows(students)
                    }
                }.readFirstSheet { sheet ->
                    sheet.mergedRefs() shouldContainExactlyInAnyOrder listOf("A1:B1", "C1:C2")

                    sheet.text(0, 0) shouldBe "성적"
                    sheet.text(0, 2) shouldBe "비고"

                    (0..2).map { sheet.text(1, it) } shouldBe listOf("국어", "영어", "비고")

                    sheet.ctWorksheet.autoFilter.ref shouldBe "A2:C2"
                    sheet.paneInformation.horizontalSplitTopRow.toInt() shouldBe 2
                }
            }

            it("한 컬럼짜리 그룹은 병합하지 않고 제목만 쓴다") {
                excel {
                    sheet<Student>("성적표") {
                        column("번호") { it.no }
                        headerGroup("국어") {
                            column("점수") { it.korean }
                        }
                        rows(students)
                    }
                }.readFirstSheet { sheet ->
                    sheet.mergedRefs() shouldContainExactlyInAnyOrder listOf("A1:A2")
                    sheet.text(0, 1) shouldBe "국어"
                    sheet.text(1, 1) shouldBe "점수"
                }
            }

            it("그룹 없는 컬럼의 위쪽 셀에는 아래쪽 셀과 같은 컬럼 헤더 스타일이 적용된다") {
                excel {
                    styles {
                        header {
                            backgroundColor(Color.GRAY)
                            bold()
                        }
                    }
                    sheet<Student>("성적표") {
                        column("번호", headerStyle = { italic() }) { it.no }
                        headerGroup("성적") {
                            column("국어") { it.korean }
                        }
                        rows(students)
                    }
                }.readFirstSheet { sheet ->
                    val top = sheet.getRow(0).getCell(0)
                    val bottom = sheet.getRow(1).getCell(0)
                    top.cellStyle.font.italic shouldBe true
                    top.cellStyle.index shouldBe bottom.cellStyle.index

                    // 그룹 제목에는 전역 헤더 스타일이 적용된다
                    val groupTitle = sheet.getRow(0).getCell(1)
                    groupTitle.cellStyle.fillPattern shouldBe FillPatternType.SOLID_FOREGROUND
                    groupTitle.cellStyle.font.italic shouldBe false
                }
            }
        }
    })
