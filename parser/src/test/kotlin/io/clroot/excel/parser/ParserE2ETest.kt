package io.clroot.excel.parser

import io.clroot.excel.annotation.Column
import io.clroot.excel.annotation.Excel
import io.clroot.excel.annotation.excelOf
import io.clroot.excel.render.writeTo
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.collections.shouldHaveSize
import io.kotest.matchers.shouldBe
import io.kotest.matchers.types.shouldBeInstanceOf
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import java.math.BigDecimal

class ParserE2ETest : DescribeSpec({

    describe("헤더 alias 테스트") {
        it("alias로 헤더를 매칭한다") {
            @Excel
            data class User(
                @Column("이름", aliases = ["Name", "성명"]) val name: String,
            )

            // 헤더가 "Name"인 엑셀 파일 생성
            val output = ByteArrayOutputStream()
            XSSFWorkbook().use { workbook ->
                val sheet = workbook.createSheet("Sheet1")
                val headerRow = sheet.createRow(0)
                headerRow.createCell(0).setCellValue("Name") // alias 사용

                val dataRow = sheet.createRow(1)
                dataRow.createCell(0).setCellValue("김철수")

                workbook.write(output)
            }

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray()))

            result.shouldBeInstanceOf<ParseResult.Success<User>>()
            result.getOrThrow()[0].name shouldBe "김철수"
        }

        it("여러 alias 중 하나로 매칭한다") {
            @Excel
            data class User(
                @Column("이름", aliases = ["Name", "성명", "Full Name"]) val name: String,
            )

            val output = ByteArrayOutputStream()
            XSSFWorkbook().use { workbook ->
                val sheet = workbook.createSheet("Sheet1")
                val headerRow = sheet.createRow(0)
                headerRow.createCell(0).setCellValue("성명") // 두 번째 alias 사용

                val dataRow = sheet.createRow(1)
                dataRow.createCell(0).setCellValue("이영희")

                workbook.write(output)
            }

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray()))

            result.shouldBeInstanceOf<ParseResult.Success<User>>()
            result.getOrThrow()[0].name shouldBe "이영희"
        }
    }

    describe("전체 검증 테스트") {
        it("validateAll이 실패하면 에러를 반환한다") {
            @Excel
            data class User(
                @Column("이름", order = 1) val name: String,
                @Column("이메일", order = 2) val email: String,
            )

            val original = listOf(
                User("김철수", "same@email.com"),
                User("이영희", "same@email.com"),
            )

            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                validateAll { users ->
                    val duplicates = users.groupBy { it.email }.filter { it.value.size > 1 }
                    require(duplicates.isEmpty()) { "중복된 이메일이 있습니다: ${duplicates.keys}" }
                }
            }

            result.shouldBeInstanceOf<ParseResult.Failure<User>>()
        }

        it("validateAll이 성공하면 데이터를 반환한다") {
            @Excel
            data class User(
                @Column("이름", order = 1) val name: String,
                @Column("이메일", order = 2) val email: String,
            )

            val original = listOf(
                User("김철수", "kim@email.com"),
                User("이영희", "lee@email.com"),
            )

            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                validateAll { users ->
                    val duplicates = users.groupBy { it.email }.filter { it.value.size > 1 }
                    require(duplicates.isEmpty()) { "중복된 이메일이 있습니다: ${duplicates.keys}" }
                }
            }

            result.shouldBeInstanceOf<ParseResult.Success<User>>()
            result.getOrThrow() shouldHaveSize 2
        }
    }

    describe("커스텀 컨버터 테스트") {
        it("커스텀 타입을 변환한다") {
            data class Money(val amount: BigDecimal)

            @Excel
            data class Product(
                @Column("상품명", order = 1) val name: String,
                @Column("가격", order = 2) val price: Money,
            )

            // 가격이 숫자로 저장된 엑셀 생성
            val output = ByteArrayOutputStream()
            XSSFWorkbook().use { workbook ->
                val sheet = workbook.createSheet("Sheet1")
                val headerRow = sheet.createRow(0)
                headerRow.createCell(0).setCellValue("상품명")
                headerRow.createCell(1).setCellValue("가격")

                val dataRow = sheet.createRow(1)
                dataRow.createCell(0).setCellValue("노트북")
                dataRow.createCell(1).setCellValue(1500000.0)

                workbook.write(output)
            }

            val result = parseExcel<Product>(ByteArrayInputStream(output.toByteArray())) {
                converter<Money> { value ->
                    Money(BigDecimal(value?.toString() ?: "0"))
                }
            }

            result.shouldBeInstanceOf<ParseResult.Success<Product>>()
            result.getOrThrow()[0].price.amount.compareTo(BigDecimal("1500000")) shouldBe 0
        }

        it("문자열을 커스텀 타입으로 변환한다") {
            data class Email(val value: String)

            @Excel
            data class Contact(
                @Column("이름", order = 1) val name: String,
                @Column("이메일", order = 2) val email: Email,
            )

            val output = ByteArrayOutputStream()
            XSSFWorkbook().use { workbook ->
                val sheet = workbook.createSheet("Sheet1")
                val headerRow = sheet.createRow(0)
                headerRow.createCell(0).setCellValue("이름")
                headerRow.createCell(1).setCellValue("이메일")

                val dataRow = sheet.createRow(1)
                dataRow.createCell(0).setCellValue("김철수")
                dataRow.createCell(1).setCellValue("kim@example.com")

                workbook.write(output)
            }

            val result = parseExcel<Contact>(ByteArrayInputStream(output.toByteArray())) {
                converter<Email> { value ->
                    Email(value?.toString() ?: "")
                }
            }

            result.shouldBeInstanceOf<ParseResult.Success<Contact>>()
            result.getOrThrow()[0].email shouldBe Email("kim@example.com")
        }
    }

    describe("시트 선택 테스트") {
        it("sheetName으로 시트를 선택한다") {
            @Excel
            data class User(
                @Column("이름") val name: String,
            )

            val output = ByteArrayOutputStream()
            XSSFWorkbook().use { workbook ->
                // 빈 시트
                workbook.createSheet("Empty")

                // 데이터가 있는 시트
                val sheet = workbook.createSheet("Users")
                sheet.createRow(0).createCell(0).setCellValue("이름")
                sheet.createRow(1).createCell(0).setCellValue("김철수")

                workbook.write(output)
            }

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                sheetName = "Users"
            }

            result.shouldBeInstanceOf<ParseResult.Success<User>>()
            result.getOrThrow()[0].name shouldBe "김철수"
        }

        it("sheetIndex로 시트를 선택한다") {
            @Excel
            data class User(
                @Column("이름") val name: String,
            )

            val output = ByteArrayOutputStream()
            XSSFWorkbook().use { workbook ->
                // 첫 번째 시트 (index 0)
                workbook.createSheet("Empty")

                // 두 번째 시트 (index 1)
                val sheet = workbook.createSheet("Users")
                sheet.createRow(0).createCell(0).setCellValue("이름")
                sheet.createRow(1).createCell(0).setCellValue("김철수")

                workbook.write(output)
            }

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                sheetIndex = 1
            }

            result.shouldBeInstanceOf<ParseResult.Success<User>>()
            result.getOrThrow()[0].name shouldBe "김철수"
        }
    }

    describe("FAIL_FAST 모드 테스트") {
        it("첫 에러에서 중단한다") {
            @Excel
            data class User(
                @Column("이름", order = 1) val name: String,
                @Column("이메일", order = 2) val email: String,
            )

            val original = listOf(
                User("김철수", "invalid1"),
                User("이영희", "invalid2"),
                User("박민수", "invalid3"),
            )

            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                onError = OnError.FAIL_FAST
                validateRow { user ->
                    require(user.email.contains("@")) { "이메일 형식 오류" }
                }
            }

            result.shouldBeInstanceOf<ParseResult.Failure<User>>()
            val errors = (result as ParseResult.Failure).errors
            errors shouldHaveSize 1 // 첫 번째 에러에서 중단
        }

        it("COLLECT 모드는 모든 에러를 수집한다") {
            @Excel
            data class User(
                @Column("이름", order = 1) val name: String,
                @Column("이메일", order = 2) val email: String,
            )

            val original = listOf(
                User("김철수", "invalid1"),
                User("이영희", "invalid2"),
                User("박민수", "invalid3"),
            )

            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                onError = OnError.COLLECT
                validateRow { user ->
                    require(user.email.contains("@")) { "이메일 형식 오류" }
                }
            }

            result.shouldBeInstanceOf<ParseResult.Failure<User>>()
            val errors = (result as ParseResult.Failure).errors
            errors shouldHaveSize 3 // 모든 에러 수집
        }
    }

    describe("헤더 매칭 전략 테스트") {
        it("FLEXIBLE 모드는 대소문자를 무시한다") {
            @Excel
            data class User(
                @Column("Name") val name: String,
            )

            val output = ByteArrayOutputStream()
            XSSFWorkbook().use { workbook ->
                val sheet = workbook.createSheet("Sheet1")
                sheet.createRow(0).createCell(0).setCellValue("name") // 소문자

                val dataRow = sheet.createRow(1)
                dataRow.createCell(0).setCellValue("김철수")

                workbook.write(output)
            }

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                headerMatching = HeaderMatching.FLEXIBLE
            }

            result.shouldBeInstanceOf<ParseResult.Success<User>>()
            result.getOrThrow()[0].name shouldBe "김철수"
        }

        it("FLEXIBLE 모드는 공백을 정규화한다") {
            @Excel
            data class User(
                @Column("Full Name") val name: String,
            )

            val output = ByteArrayOutputStream()
            XSSFWorkbook().use { workbook ->
                val sheet = workbook.createSheet("Sheet1")
                sheet.createRow(0).createCell(0).setCellValue("  Full   Name  ") // 여러 공백

                val dataRow = sheet.createRow(1)
                dataRow.createCell(0).setCellValue("김철수")

                workbook.write(output)
            }

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                headerMatching = HeaderMatching.FLEXIBLE
            }

            result.shouldBeInstanceOf<ParseResult.Success<User>>()
            result.getOrThrow()[0].name shouldBe "김철수"
        }

        it("EXACT 모드는 정확히 일치해야 한다") {
            @Excel
            data class User(
                @Column("Name") val name: String,
            )

            val output = ByteArrayOutputStream()
            XSSFWorkbook().use { workbook ->
                val sheet = workbook.createSheet("Sheet1")
                sheet.createRow(0).createCell(0).setCellValue("name") // 소문자

                val dataRow = sheet.createRow(1)
                dataRow.createCell(0).setCellValue("김철수")

                workbook.write(output)
            }

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                headerMatching = HeaderMatching.EXACT
            }

            // EXACT 모드에서는 "Name"과 "name"이 매칭되지 않아 필수 필드 누락
            result.shouldBeInstanceOf<ParseResult.Failure<User>>()
        }
    }

    describe("빈 행 처리 테스트") {
        it("skipEmptyRows가 true면 빈 행을 건너뛴다") {
            @Excel
            data class User(
                @Column("이름") val name: String,
            )

            val output = ByteArrayOutputStream()
            XSSFWorkbook().use { workbook ->
                val sheet = workbook.createSheet("Sheet1")
                sheet.createRow(0).createCell(0).setCellValue("이름")
                sheet.createRow(1).createCell(0).setCellValue("김철수")
                sheet.createRow(2) // 빈 행
                sheet.createRow(3).createCell(0).setCellValue("이영희")

                workbook.write(output)
            }

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                skipEmptyRows = true
            }

            result.shouldBeInstanceOf<ParseResult.Success<User>>()
            val users = result.getOrThrow()
            users shouldHaveSize 2
            users[0].name shouldBe "김철수"
            users[1].name shouldBe "이영희"
        }
    }

    describe("라운드트립 테스트") {
        it("쓰기와 읽기의 결과가 동일하다") {
            @Excel
            data class Employee(
                @Column("사번", order = 1) val id: Int,
                @Column("이름", order = 2) val name: String,
                @Column("부서", order = 3) val department: String,
                @Column("연봉", order = 4) val salary: Long,
            )

            val original = listOf(
                Employee(1001, "김철수", "개발팀", 50000000),
                Employee(1002, "이영희", "마케팅팀", 45000000),
                Employee(1003, "박민수", "인사팀", 48000000),
            )

            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<Employee>(ByteArrayInputStream(output.toByteArray()))

            result.shouldBeInstanceOf<ParseResult.Success<Employee>>()
            val parsed = result.getOrThrow()
            parsed shouldHaveSize 3
            parsed shouldBe original
        }
    }
})
