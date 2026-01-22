package io.clroot.excel.parser

import io.clroot.excel.annotation.Column
import io.clroot.excel.annotation.Excel
import io.clroot.excel.core.ExcelConfigurationException
import io.kotest.assertions.throwables.shouldThrow
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.collections.shouldHaveSize
import io.kotest.matchers.shouldBe
import io.kotest.matchers.string.shouldContain

class AnnotationExtractorTest : DescribeSpec({

    describe("extractColumns") {
        it("@Column 어노테이션이 있는 프로퍼티를 추출한다") {
            @Excel
            data class User(
                @Column("이름") val name: String,
                @Column("나이") val age: Int,
            )

            val columns = AnnotationExtractor.extractColumns(User::class)

            columns shouldHaveSize 2
            columns[0].header shouldBe "이름"
            columns[0].propertyName shouldBe "name"
            columns[1].header shouldBe "나이"
            columns[1].propertyName shouldBe "age"
        }

        it("aliases를 추출한다") {
            @Excel
            data class User(
                @Column("이름", aliases = ["Name", "성명"]) val name: String,
            )

            val columns = AnnotationExtractor.extractColumns(User::class)

            columns[0].aliases shouldBe arrayOf("Name", "성명")
        }

        it("order 순서대로 정렬된다") {
            @Excel
            data class User(
                @Column("나이", order = 2) val age: Int,
                @Column("이름", order = 1) val name: String,
            )

            val columns = AnnotationExtractor.extractColumns(User::class)

            columns[0].header shouldBe "이름"
            columns[1].header shouldBe "나이"
        }

        it("nullable 타입을 감지한다") {
            @Excel
            data class User(
                @Column("이름") val name: String,
                @Column("별명") val nickname: String?,
            )

            val columns = AnnotationExtractor.extractColumns(User::class)

            columns.find { it.propertyName == "name" }?.isNullable shouldBe false
            columns.find { it.propertyName == "nickname" }?.isNullable shouldBe true
        }

        it("@Excel 어노테이션이 없으면 예외를 던진다") {
            data class User(
                @Column("이름") val name: String,
            )

            val exception =
                shouldThrow<ExcelConfigurationException> {
                    AnnotationExtractor.extractColumns(User::class)
                }
            exception.message shouldContain "@Excel"
        }

        it("@Column 어노테이션이 없으면 예외를 던진다") {
            @Excel
            data class User(val name: String)

            val exception =
                shouldThrow<ExcelConfigurationException> {
                    AnnotationExtractor.extractColumns(User::class)
                }
            exception.message shouldContain "@Column"
        }
    }
})
