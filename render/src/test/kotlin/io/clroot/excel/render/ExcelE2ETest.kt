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
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import java.time.LocalDate

class ExcelE2ETest :
    DescribeSpec({

        describe("DSL 방식") {
            it("기본 Excel 파일을 생성한다") {
                data class User(
                    val name: String,
                    val age: Int,
                )

                val users =
                    listOf(
                        User("김철수", 30),
                        User("이영희", 25),
                    )

                val document =
                    excel {
                        sheet<User>("사용자") {
                            column("이름") { it.name }
                            column("나이") { it.age }
                            rows(users)
                        }
                    }

                val output = ByteArrayOutputStream()
                document.writeTo(output)

                XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                    val sheet = workbook.getSheetAt(0)

                    sheet.sheetName shouldBe "사용자"
                    sheet.getRow(0).getCell(0).stringCellValue shouldBe "이름"
                    sheet.getRow(0).getCell(1).stringCellValue shouldBe "나이"
                    sheet.getRow(1).getCell(0).stringCellValue shouldBe "김철수"
                    sheet.getRow(1).getCell(1).numericCellValue shouldBe 30.0
                    sheet.getRow(2).getCell(0).stringCellValue shouldBe "이영희"
                    sheet.getRow(2).getCell(1).numericCellValue shouldBe 25.0
                }
            }

            it("고정 컬럼 너비를 설정한다") {
                data class User(
                    val name: String,
                    val description: String,
                )

                val users = listOf(User("김철수", "개발자"))

                val document =
                    excel {
                        sheet<User>("사용자") {
                            column("이름", width = 10.chars) { it.name }
                            column("설명", width = 30.chars) { it.description }
                            rows(users)
                        }
                    }

                val output = ByteArrayOutputStream()
                document.writeTo(output)

                XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                    val sheet = workbook.getSheetAt(0)
                    // POI에서 컬럼 너비는 1/256 문자 단위
                    sheet.getColumnWidth(0) shouldBe 10 * 256
                    sheet.getColumnWidth(1) shouldBe 30 * 256
                }
            }

            it("멀티 시트를 생성한다") {
                data class Product(
                    val name: String,
                    val price: Int,
                )

                data class Order(
                    val id: String,
                    val amount: Int,
                )

                val products = listOf(Product("노트북", 1500000), Product("마우스", 50000))
                val orders = listOf(Order("ORD-001", 2), Order("ORD-002", 5))

                val document =
                    excel {
                        sheet<Product>("상품") {
                            column("상품명") { it.name }
                            column("가격") { it.price }
                            rows(products)
                        }
                        sheet<Order>("주문") {
                            column("주문번호") { it.id }
                            column("수량") { it.amount }
                            rows(orders)
                        }
                    }

                val output = ByteArrayOutputStream()
                document.writeTo(output)

                XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                    workbook.numberOfSheets shouldBe 2
                    workbook.getSheetAt(0).sheetName shouldBe "상품"
                    workbook.getSheetAt(1).sheetName shouldBe "주문"

                    val productSheet = workbook.getSheetAt(0)
                    productSheet.getRow(1).getCell(0).stringCellValue shouldBe "노트북"
                    productSheet.getRow(1).getCell(1).numericCellValue shouldBe 1500000.0

                    val orderSheet = workbook.getSheetAt(1)
                    orderSheet.getRow(1).getCell(0).stringCellValue shouldBe "ORD-001"
                    orderSheet.getRow(1).getCell(1).numericCellValue shouldBe 2.0
                }
            }
        }

        describe("스타일링") {
            it("컬럼별 인라인 스타일을 적용한다") {
                data class Product(
                    val name: String,
                    val price: Int,
                )

                val products = listOf(Product("노트북", 1500000))

                val document =
                    excel {
                        sheet<Product>("상품") {
                            column("상품명") { it.name }
                            column("가격", style = {
                                align(Alignment.RIGHT)
                                bold()
                            }) { it.price }
                            rows(products)
                        }
                    }

                val output = ByteArrayOutputStream()
                document.writeTo(output)

                XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                    val sheet = workbook.getSheetAt(0)
                    val priceDataCell = sheet.getRow(1).getCell(1)

                    priceDataCell.cellStyle.alignment shouldBe HorizontalAlignment.RIGHT
                    priceDataCell.cellStyle.font.bold shouldBe true
                }
            }

            it("styles DSL에서 컬럼별 스타일을 적용한다") {
                data class Product(
                    val name: String,
                    val price: Int,
                )

                val products = listOf(Product("노트북", 1500000))

                val document =
                    excel {
                        styles {
                            header {
                                backgroundColor(Color.GRAY)
                                bold()
                            }
                            column("가격") {
                                body {
                                    align(Alignment.RIGHT)
                                }
                            }
                        }
                        sheet<Product>("상품") {
                            column("상품명") { it.name }
                            column("가격") { it.price }
                            rows(products)
                        }
                    }

                val output = ByteArrayOutputStream()
                document.writeTo(output)

                XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                    val sheet = workbook.getSheetAt(0)

                    // 헤더는 전역 스타일 적용
                    val headerCell = sheet.getRow(0).getCell(0)
                    headerCell.cellStyle.fillPattern shouldBe FillPatternType.SOLID_FOREGROUND
                    headerCell.cellStyle.font.bold shouldBe true

                    // 가격 컬럼은 오른쪽 정렬
                    val priceDataCell = sheet.getRow(1).getCell(1)
                    priceDataCell.cellStyle.alignment shouldBe HorizontalAlignment.RIGHT
                }
            }

            it("테마를 적용하면 헤더에 스타일이 적용된다") {
                data class User(
                    val name: String,
                    val age: Int,
                )

                val users = listOf(User("김철수", 30))

                val document =
                    excel(theme = Theme.Modern) {
                        sheet<User>("사용자") {
                            column("이름") { it.name }
                            column("나이") { it.age }
                            rows(users)
                        }
                    }

                val output = ByteArrayOutputStream()
                document.writeTo(output)

                XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                    val sheet = workbook.getSheetAt(0)
                    val headerCell = sheet.getRow(0).getCell(0)
                    val headerStyle = headerCell.cellStyle

                    // Modern 테마: 파란 배경, 흰색 글자, 볼드, 가운데 정렬
                    headerStyle.fillPattern shouldBe FillPatternType.SOLID_FOREGROUND
                    headerStyle.font.bold shouldBe true
                    headerStyle.alignment shouldBe HorizontalAlignment.CENTER
                }
            }

            it("styles DSL로 헤더/본문 스타일을 설정한다") {
                data class User(
                    val name: String,
                    val age: Int,
                )

                val users = listOf(User("김철수", 30))

                val document =
                    excel {
                        styles {
                            header {
                                backgroundColor(Color.GRAY)
                                bold()
                                align(Alignment.CENTER)
                            }
                        }
                        sheet<User>("사용자") {
                            column("이름") { it.name }
                            column("나이") { it.age }
                            rows(users)
                        }
                    }

                val output = ByteArrayOutputStream()
                document.writeTo(output)

                XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                    val sheet = workbook.getSheetAt(0)
                    val headerCell = sheet.getRow(0).getCell(0)

                    headerCell.cellStyle.fillPattern shouldBe FillPatternType.SOLID_FOREGROUND
                    headerCell.cellStyle.font.bold shouldBe true
                    headerCell.cellStyle.alignment shouldBe HorizontalAlignment.CENTER
                }
            }
        }

        describe("헤더 그룹") {
            it("2단 헤더를 생성한다") {
                data class Student(
                    val name: String,
                    val studentId: String,
                    val korean: Int,
                    val english: Int,
                    val math: Int,
                )

                val students =
                    listOf(
                        Student("김철수", "20241", 90, 85, 95),
                        Student("이영희", "20242", 88, 92, 87),
                    )

                val document =
                    excel {
                        sheet<Student>("성적표") {
                            headerGroup("학생 정보") {
                                column("이름") { it.name }
                                column("학번") { it.studentId }
                            }
                            headerGroup("성적") {
                                column("국어") { it.korean }
                                column("영어") { it.english }
                                column("수학") { it.math }
                            }
                            rows(students)
                        }
                    }

                val output = ByteArrayOutputStream()
                document.writeTo(output)

                XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                    val sheet = workbook.getSheetAt(0)

                    // 첫 번째 행: 그룹 헤더
                    sheet.getRow(0).getCell(0).stringCellValue shouldBe "학생 정보"
                    sheet.getRow(0).getCell(2).stringCellValue shouldBe "성적"

                    // 두 번째 행: 컬럼 헤더
                    sheet.getRow(1).getCell(0).stringCellValue shouldBe "이름"
                    sheet.getRow(1).getCell(1).stringCellValue shouldBe "학번"
                    sheet.getRow(1).getCell(2).stringCellValue shouldBe "국어"
                    sheet.getRow(1).getCell(3).stringCellValue shouldBe "영어"
                    sheet.getRow(1).getCell(4).stringCellValue shouldBe "수학"

                    // 데이터 행 (3번째 행부터)
                    sheet.getRow(2).getCell(0).stringCellValue shouldBe "김철수"
                    sheet.getRow(2).getCell(2).numericCellValue shouldBe 90.0

                    // 셀 병합 확인 (학생 정보: 0-1, 성적: 2-4)
                    val mergedRegions = sheet.mergedRegions
                    mergedRegions.size shouldBe 2
                }
            }
        }

        describe("날짜/시간 포맷") {
            it("LocalDate는 자동으로 날짜 형식으로 렌더링된다") {
                data class Event(
                    val name: String,
                    val date: LocalDate,
                )

                val events = listOf(Event("회의", LocalDate.of(2024, 5, 15)))

                val document =
                    excel {
                        sheet<Event>("일정") {
                            column("일정명") { it.name }
                            column("날짜") { it.date }
                            rows(events)
                        }
                    }

                val output = ByteArrayOutputStream()
                document.writeTo(output)

                XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                    val sheet = workbook.getSheetAt(0)
                    val dateCell = sheet.getRow(1).getCell(1)

                    // 날짜가 Excel 날짜 형식으로 저장되어야 함
                    dateCell.localDateTimeCellValue.toLocalDate() shouldBe LocalDate.of(2024, 5, 15)
                }
            }
        }

        describe("어노테이션 방식") {
            it("@Excel, @Column으로 Excel 파일을 생성한다") {
                val users =
                    listOf(
                        AnnotatedUser("김철수", 30, LocalDate.of(2024, 1, 15)),
                        AnnotatedUser("이영희", 25, LocalDate.of(2024, 3, 20)),
                    )

                val document = excelOf(users)

                val output = ByteArrayOutputStream()
                document.writeTo(output)

                XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                    val sheet = workbook.getSheetAt(0)

                    // order 속성으로 정렬되므로 순서 보장
                    sheet.getRow(0).getCell(0).stringCellValue shouldBe "이름"
                    sheet.getRow(0).getCell(1).stringCellValue shouldBe "나이"
                    sheet.getRow(0).getCell(2).stringCellValue shouldBe "가입일"
                    sheet.getRow(1).getCell(0).stringCellValue shouldBe "김철수"
                    sheet.getRow(1).getCell(1).numericCellValue shouldBe 30.0
                }
            }

            it("excelOf에 테마를 적용하면 헤더에 스타일이 적용된다") {
                val users =
                    listOf(
                        AnnotatedUser("김철수", 30, LocalDate.of(2024, 1, 15)),
                    )

                val document = excelOf(users, theme = Theme.Modern)

                val output = ByteArrayOutputStream()
                document.writeTo(output)

                XSSFWorkbook(ByteArrayInputStream(output.toByteArray())).use { workbook ->
                    val sheet = workbook.getSheetAt(0)
                    val headerCell = sheet.getRow(0).getCell(0)
                    val headerStyle = headerCell.cellStyle

                    // Modern 테마: 파란 배경, 흰색 글자, 볼드, 가운데 정렬
                    headerStyle.fillPattern shouldBe FillPatternType.SOLID_FOREGROUND
                    headerStyle.font.bold shouldBe true
                    headerStyle.alignment shouldBe HorizontalAlignment.CENTER

                    // Modern 테마 Body: 얇은 테두리
                    val bodyCell = sheet.getRow(1).getCell(0)
                    val bodyStyle = bodyCell.cellStyle
                    bodyStyle.borderBottom shouldBe org.apache.poi.ss.usermodel.BorderStyle.THIN
                }
            }
        }
    })

@Excel
data class AnnotatedUser(
    @Column("이름", order = 1)
    val name: String,
    @Column("나이", order = 2)
    val age: Int,
    @Column("가입일", order = 3)
    val joinedAt: LocalDate,
)
