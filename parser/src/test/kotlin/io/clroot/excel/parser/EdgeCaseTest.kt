package io.clroot.excel.parser

import io.clroot.excel.annotation.Column
import io.clroot.excel.annotation.Excel
import io.clroot.excel.annotation.excelOf
import io.clroot.excel.render.writeTo
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.collections.shouldBeEmpty
import io.kotest.matchers.collections.shouldHaveSize
import io.kotest.matchers.shouldBe
import io.kotest.matchers.types.shouldBeInstanceOf
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream

/**
 * Edge case tests for Excel parsing.
 *
 * Tests various edge cases including:
 * - Empty sheets
 * - Special characters
 * - Unicode content
 * - Missing columns
 * - Extra columns in Excel
 */
class EdgeCaseTest :
    DescribeSpec({

        describe("빈 데이터 처리") {
            it("빈 데이터 리스트를 파싱하면 빈 결과를 반환한다") {
                val original = emptyList<SimpleUser>()

                val output = ByteArrayOutputStream()
                excelOf(original).writeTo(output)

                val result = parseExcel<SimpleUser>(ByteArrayInputStream(output.toByteArray()))

                result.shouldBeInstanceOf<ParseResult.Success<SimpleUser>>()
                result.getOrThrow().shouldBeEmpty()
            }

            it("헤더만 있는 시트를 파싱한다") {
                // Create Excel with header only
                val output = ByteArrayOutputStream()
                XSSFWorkbook().use { workbook ->
                    val sheet = workbook.createSheet("Sheet1")
                    val headerRow = sheet.createRow(0)
                    headerRow.createCell(0).setCellValue("이름")
                    headerRow.createCell(1).setCellValue("나이")
                    workbook.write(output)
                }

                val result = parseExcel<SimpleUser>(ByteArrayInputStream(output.toByteArray()))

                result.shouldBeInstanceOf<ParseResult.Success<SimpleUser>>()
                result.getOrThrow().shouldBeEmpty()
            }
        }

        describe("특수 문자 처리") {
            it("이모지를 포함한 데이터를 파싱한다") {
                val original = listOf(SimpleUser("김철수 😀", 30))

                val output = ByteArrayOutputStream()
                excelOf(original).writeTo(output)

                val result = parseExcel<SimpleUser>(ByteArrayInputStream(output.toByteArray()))

                result.shouldBeInstanceOf<ParseResult.Success<SimpleUser>>()
                val users = result.getOrThrow()
                users shouldHaveSize 1
                users[0].name shouldBe "김철수 😀"
            }

            it("개행 문자를 포함한 데이터를 파싱한다") {
                val original = listOf(SimpleUser("김철수\n(개발팀)", 30))

                val output = ByteArrayOutputStream()
                excelOf(original).writeTo(output)

                val result = parseExcel<SimpleUser>(ByteArrayInputStream(output.toByteArray()))

                result.shouldBeInstanceOf<ParseResult.Success<SimpleUser>>()
                result.getOrThrow()[0].name shouldBe "김철수\n(개발팀)"
            }

            it("탭 문자를 포함한 데이터를 파싱한다") {
                val original = listOf(SimpleUser("김철수\t개발팀", 30))

                val output = ByteArrayOutputStream()
                excelOf(original).writeTo(output)

                val result = parseExcel<SimpleUser>(ByteArrayInputStream(output.toByteArray()))

                result.shouldBeInstanceOf<ParseResult.Success<SimpleUser>>()
                result.getOrThrow()[0].name shouldBe "김철수\t개발팀"
            }

            it("따옴표를 포함한 데이터를 파싱한다") {
                val original = listOf(SimpleUser("김철수 \"닉네임\"", 30))

                val output = ByteArrayOutputStream()
                excelOf(original).writeTo(output)

                val result = parseExcel<SimpleUser>(ByteArrayInputStream(output.toByteArray()))

                result.shouldBeInstanceOf<ParseResult.Success<SimpleUser>>()
                result.getOrThrow()[0].name shouldBe "김철수 \"닉네임\""
            }
        }

        describe("다국어 처리") {
            it("한글 데이터를 파싱한다") {
                val original = listOf(SimpleUser("김철수", 30))

                val output = ByteArrayOutputStream()
                excelOf(original).writeTo(output)

                val result = parseExcel<SimpleUser>(ByteArrayInputStream(output.toByteArray()))

                result.shouldBeInstanceOf<ParseResult.Success<SimpleUser>>()
                result.getOrThrow()[0].name shouldBe "김철수"
            }

            it("일본어 데이터를 파싱한다") {
                val original = listOf(SimpleUser("田中太郎", 30))

                val output = ByteArrayOutputStream()
                excelOf(original).writeTo(output)

                val result = parseExcel<SimpleUser>(ByteArrayInputStream(output.toByteArray()))

                result.shouldBeInstanceOf<ParseResult.Success<SimpleUser>>()
                result.getOrThrow()[0].name shouldBe "田中太郎"
            }

            it("중국어 데이터를 파싱한다") {
                val original = listOf(SimpleUser("王小明", 30))

                val output = ByteArrayOutputStream()
                excelOf(original).writeTo(output)

                val result = parseExcel<SimpleUser>(ByteArrayInputStream(output.toByteArray()))

                result.shouldBeInstanceOf<ParseResult.Success<SimpleUser>>()
                result.getOrThrow()[0].name shouldBe "王小明"
            }

            it("아랍어 데이터를 파싱한다") {
                val original = listOf(SimpleUser("محمد", 30))

                val output = ByteArrayOutputStream()
                excelOf(original).writeTo(output)

                val result = parseExcel<SimpleUser>(ByteArrayInputStream(output.toByteArray()))

                result.shouldBeInstanceOf<ParseResult.Success<SimpleUser>>()
                result.getOrThrow()[0].name shouldBe "محمد"
            }
        }

        describe("숫자 경계값 처리") {
            it("큰 정수를 파싱한다") {
                val original = listOf(SimpleUser("사용자", Int.MAX_VALUE))

                val output = ByteArrayOutputStream()
                excelOf(original).writeTo(output)

                val result = parseExcel<SimpleUser>(ByteArrayInputStream(output.toByteArray()))

                result.shouldBeInstanceOf<ParseResult.Success<SimpleUser>>()
                result.getOrThrow()[0].age shouldBe Int.MAX_VALUE
            }

            it("0을 파싱한다") {
                val original = listOf(SimpleUser("사용자", 0))

                val output = ByteArrayOutputStream()
                excelOf(original).writeTo(output)

                val result = parseExcel<SimpleUser>(ByteArrayInputStream(output.toByteArray()))

                result.shouldBeInstanceOf<ParseResult.Success<SimpleUser>>()
                result.getOrThrow()[0].age shouldBe 0
            }

            it("음수를 파싱한다") {
                val original = listOf(SimpleUser("사용자", -100))

                val output = ByteArrayOutputStream()
                excelOf(original).writeTo(output)

                val result = parseExcel<SimpleUser>(ByteArrayInputStream(output.toByteArray()))

                result.shouldBeInstanceOf<ParseResult.Success<SimpleUser>>()
                result.getOrThrow()[0].age shouldBe -100
            }
        }

        describe("nullable 필드 처리") {
            it("null 값이 있는 nullable 필드를 파싱한다") {
                val original =
                    listOf(
                        UserWithNullable("김철수", "닉네임"),
                        UserWithNullable("이영희", null),
                    )

                val output = ByteArrayOutputStream()
                excelOf(original).writeTo(output)

                val result = parseExcel<UserWithNullable>(ByteArrayInputStream(output.toByteArray()))

                result.shouldBeInstanceOf<ParseResult.Success<UserWithNullable>>()
                val users = result.getOrThrow()
                users shouldHaveSize 2
                users[0].nickname shouldBe "닉네임"
                users[1].nickname shouldBe null
            }
        }

        describe("빈 문자열 처리") {
            it("빈 문자열이 있는 nullable 필드를 파싱한다") {
                val original = listOf(UserWithNullable("김철수", ""))

                val output = ByteArrayOutputStream()
                excelOf(original).writeTo(output)

                val result =
                    parseExcel<UserWithNullable>(ByteArrayInputStream(output.toByteArray())) {
                        treatBlankAsNull = false
                    }

                result.shouldBeInstanceOf<ParseResult.Success<UserWithNullable>>()
                result.getOrThrow()[0].nickname shouldBe ""
            }

            it("treatBlankAsNull=true이면 빈 문자열을 null로 처리한다") {
                val original = listOf(UserWithNullable("김철수", ""))

                val output = ByteArrayOutputStream()
                excelOf(original).writeTo(output)

                val result =
                    parseExcel<UserWithNullable>(ByteArrayInputStream(output.toByteArray())) {
                        treatBlankAsNull = true
                    }

                result.shouldBeInstanceOf<ParseResult.Success<UserWithNullable>>()
                result.getOrThrow()[0].nickname shouldBe null
            }
        }

        describe("대용량 문자열 처리") {
            it("긴 문자열을 파싱한다") {
                val longName = "가".repeat(10000)
                val original = listOf(SimpleUser(longName, 30))

                val output = ByteArrayOutputStream()
                excelOf(original).writeTo(output)

                val result = parseExcel<SimpleUser>(ByteArrayInputStream(output.toByteArray()))

                result.shouldBeInstanceOf<ParseResult.Success<SimpleUser>>()
                result.getOrThrow()[0].name shouldBe longName
            }
        }

        describe("Excel에 추가 컬럼이 있는 경우") {
            it("정의되지 않은 추가 컬럼을 무시한다") {
                val output = ByteArrayOutputStream()
                XSSFWorkbook().use { workbook ->
                    val sheet = workbook.createSheet("Sheet1")

                    // Header with extra column
                    val headerRow = sheet.createRow(0)
                    headerRow.createCell(0).setCellValue("이름")
                    headerRow.createCell(1).setCellValue("나이")
                    headerRow.createCell(2).setCellValue("추가컬럼") // Extra column

                    // Data with extra column
                    val dataRow = sheet.createRow(1)
                    dataRow.createCell(0).setCellValue("김철수")
                    dataRow.createCell(1).setCellValue(30.0)
                    dataRow.createCell(2).setCellValue("무시될 값")

                    workbook.write(output)
                }

                val result = parseExcel<SimpleUser>(ByteArrayInputStream(output.toByteArray()))

                result.shouldBeInstanceOf<ParseResult.Success<SimpleUser>>()
                val users = result.getOrThrow()
                users shouldHaveSize 1
                users[0].name shouldBe "김철수"
                users[0].age shouldBe 30
            }
        }
    })

@Excel
data class SimpleUser(
    @Column("이름", order = 1)
    val name: String,
    @Column("나이", order = 2)
    val age: Int,
)

@Excel
data class UserWithNullable(
    @Column("이름", order = 1)
    val name: String,
    @Column("별명", order = 2)
    val nickname: String?,
)
