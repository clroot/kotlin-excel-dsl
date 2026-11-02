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

    describe("정수와 Boolean Excel 변환") {
        @Excel
        data class Raw(
            @Column("value") val value: String,
        )

        @Excel
        data class IntegerRow(
            @Column("value") val value: Long,
        )

        @Excel
        data class BooleanRow(
            @Column("value") val value: Boolean,
        )

        fun workbook(vararg values: String): ByteArray {
            val output = ByteArrayOutputStream()
            excelOf(values.map { Raw(it) }).writeTo(output)
            return output.toByteArray()
        }

        it("텍스트 셀의 큰 정수를 정확히 읽는다") {
            val result = parseExcel<IntegerRow>(ByteArrayInputStream(workbook("9007199254740993")))
            result.getOrThrow().single().value shouldBe 9007199254740993L
        }

        it("정수 변환 오류의 행과 열을 정책에 맞게 보고한다") {
            val bytes = workbook("9007199254740993", "1.9", "9223372036854775808", "NaN")
            for (policy in listOf(OnError.FAIL_FAST, OnError.COLLECT)) {
                val result = parseExcel<IntegerRow>(ByteArrayInputStream(bytes)) { onError = policy }
                result.shouldBeInstanceOf<ParseResult.Failure<IntegerRow>>()
                val errors = (result as ParseResult.Failure).errors
                errors.map { it.rowIndex } shouldBe if (policy == OnError.FAIL_FAST) listOf(2) else listOf(2, 3, 4)
                errors.forEach {
                    it.columnHeader shouldBe "value"
                    it.cause.shouldBeInstanceOf<IllegalArgumentException>()
                }
            }
        }

        it("잘못된 Boolean 입력을 정책에 맞게 보고한다") {
            for (policy in listOf(OnError.FAIL_FAST, OnError.COLLECT)) {
                val result =
                    parseExcel<BooleanRow>(ByteArrayInputStream(workbook("true", "typo", "2"))) {
                        onError = policy
                    }
                result.shouldBeInstanceOf<ParseResult.Failure<BooleanRow>>()
                (result as ParseResult.Failure).errors.map { it.rowIndex } shouldBe
                    if (policy == OnError.FAIL_FAST) listOf(2) else listOf(2, 3)
            }
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
