package io.clroot.excel.parser

import io.clroot.excel.annotation.Column
import io.clroot.excel.annotation.Excel
import io.clroot.excel.annotation.excelOf
import io.clroot.excel.render.writeTo
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.collections.shouldHaveSize
import io.kotest.matchers.shouldBe
import io.kotest.matchers.types.shouldBeInstanceOf
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import java.time.LocalDate

class ExcelParserTest : DescribeSpec({

    describe("parseExcel 기본 사용") {
        it("기본 데이터 클래스를 파싱한다") {
            @Excel
            data class User(
                @Column("이름", order = 1) val name: String,
                @Column("나이", order = 2) val age: Int,
            )

            val original =
                listOf(
                    User("김철수", 30),
                    User("이영희", 25),
                )

            // Write to Excel
            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            // Parse back
            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray()))

            result.shouldBeInstanceOf<ParseResult.Success<User>>()
            val users = result.getOrThrow()
            users shouldHaveSize 2
            users[0] shouldBe User("김철수", 30)
            users[1] shouldBe User("이영희", 25)
        }

        it("nullable 필드를 파싱한다") {
            @Excel
            data class User(
                @Column("이름", order = 1) val name: String,
                @Column("별명", order = 2) val nickname: String?,
            )

            val original =
                listOf(
                    User("김철수", "철수"),
                    User("이영희", null),
                )

            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<User>(ByteArrayInputStream(output.toByteArray()))

            val users = result.getOrThrow()
            users[0].nickname shouldBe "철수"
            users[1].nickname shouldBe null
        }

        it("LocalDate를 파싱한다") {
            @Excel
            data class Event(
                @Column("일정명", order = 1) val name: String,
                @Column("날짜", order = 2) val date: LocalDate,
            )

            val original = listOf(Event("회의", LocalDate.of(2024, 5, 15)))

            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result = parseExcel<Event>(ByteArrayInputStream(output.toByteArray()))

            val events = result.getOrThrow()
            events[0].date shouldBe LocalDate.of(2024, 5, 15)
        }
    }

    describe("검증") {
        it("validateRow가 실패하면 에러를 수집한다") {
            @Excel
            data class User(
                @Column("이름", order = 1) val name: String,
                @Column("이메일", order = 2) val email: String,
            )

            val original =
                listOf(
                    User("김철수", "valid@email.com"),
                    User("이영희", "invalid-email"),
                )

            val output = ByteArrayOutputStream()
            excelOf(original).writeTo(output)

            val result =
                parseExcel<User>(ByteArrayInputStream(output.toByteArray())) {
                    validateRow { user ->
                        require(user.email.contains("@")) { "이메일 형식이 올바르지 않습니다" }
                    }
                }

            result.shouldBeInstanceOf<ParseResult.Failure<User>>()
            val errors = (result as ParseResult.Failure).errors
            errors shouldHaveSize 1
            errors[0].rowIndex shouldBe 2 // 0=header, 1=first data, 2=second data
        }
    }
})
