package io.clroot.excel.parser

import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe
import io.kotest.matchers.types.shouldBeInstanceOf

class ParseResultTest : DescribeSpec({

    describe("ParseResult.Success") {
        it("getOrThrow returns the data") {
            val result: ParseResult<String> = ParseResult.Success(listOf("a", "b"))

            result.getOrThrow() shouldBe listOf("a", "b")
        }

        it("getOrElse returns the data") {
            val result: ParseResult<String> = ParseResult.Success(listOf("a", "b"))

            result.getOrElse { emptyList() } shouldBe listOf("a", "b")
        }

        it("getOrNull returns the data") {
            val result: ParseResult<String> = ParseResult.Success(listOf("a", "b"))

            result.getOrNull() shouldBe listOf("a", "b")
        }
    }

    describe("ParseResult.Failure") {
        it("getOrThrow throws ExcelParseException") {
            val errors = listOf(ParseError(rowIndex = 1, columnHeader = "Name", message = "Required value missing"))
            val result: ParseResult<String> = ParseResult.Failure(errors)

            val exception = runCatching { result.getOrThrow() }.exceptionOrNull()
            exception.shouldBeInstanceOf<ExcelParseException>()
            (exception as ExcelParseException).errors shouldBe errors
        }

        it("getOrElse returns the default value") {
            val errors = listOf(ParseError(rowIndex = 1, columnHeader = "Name", message = "Required value missing"))
            val result: ParseResult<String> = ParseResult.Failure(errors)

            result.getOrElse { listOf("default") } shouldBe listOf("default")
        }

        it("getOrNull returns null") {
            val errors = listOf(ParseError(rowIndex = 1, columnHeader = "Name", message = "Required value missing"))
            val result: ParseResult<String> = ParseResult.Failure(errors)

            result.getOrNull() shouldBe null
        }
    }
})
