package io.clroot.excel.render

import io.clroot.excel.annotation.Column
import io.clroot.excel.annotation.ConditionalStyle
import io.clroot.excel.annotation.ConditionalStyler
import io.clroot.excel.annotation.Excel
import io.clroot.excel.annotation.excelOf
import io.clroot.excel.core.ExcelConfigurationException
import io.clroot.excel.core.ExcelWriteException
import io.clroot.excel.core.dsl.excel
import io.clroot.excel.core.style.CellStyle
import io.kotest.assertions.throwables.shouldThrow
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe
import io.kotest.matchers.string.shouldContain
import io.kotest.matchers.types.shouldBeSameInstanceAs
import java.io.ByteArrayOutputStream

class StringStyler : ConditionalStyler<String> {
    override fun style(value: String?): CellStyle? = if (value?.isEmpty() == true) CellStyle(bold = true) else null
}

@Excel
data class MismatchedStylerRow(
    @Column("Amount")
    @ConditionalStyle(StringStyler::class)
    val amount: Int,
)

class RenderExceptionTest : DescribeSpec({

    describe("PoiRenderer exception handling") {
        it("propagates ConditionalStyler type mismatch as ExcelConfigurationException") {
            val document = excelOf(listOf(MismatchedStylerRow(100)))

            val exception =
                shouldThrow<ExcelConfigurationException> {
                    document.writeTo(ByteArrayOutputStream())
                }

            exception.message shouldContain "ConditionalStyler type mismatch"
        }

        it("wraps value extractor failures with sheet, row and column location") {
            data class User(val id: Int, val name: String)

            val failure = IllegalStateException("boom")
            val users = (1..5).map { User(it, "user$it") }
            val document =
                excel {
                    sheet<User>("Users") {
                        column("ID") { it.id }
                        column("Name") { if (it.id == 3) throw failure else it.name }
                        rows(users)
                    }
                }

            val exception =
                shouldThrow<ExcelWriteException> {
                    document.writeTo(ByteArrayOutputStream())
                }

            // Third data row is sheet row index 3 (header occupies row 0).
            exception.sheetName shouldBe "Users"
            exception.rowIndex shouldBe 3
            exception.columnIndex shouldBe 1
            exception.message shouldContain "sheet='Users', row=4, column=2"
            exception.cause shouldBeSameInstanceAs failure
        }
    }
})
