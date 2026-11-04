package io.clroot.excel.parser

import io.clroot.excel.annotation.Column
import io.clroot.excel.annotation.Excel
import io.clroot.excel.annotation.excelOf
import io.clroot.excel.render.writeTo
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream

@Excel
data class RoundTripUser(
    @Column("Name") val name: String,
    @Column("Age") val age: Int,
)

class ColumnOrderRoundTripTest : DescribeSpec({

    describe("write and parse without @Column.order") {
        it("exports in declaration order and parses back the same data") {
            val users = listOf(RoundTripUser("Kim", 30), RoundTripUser("Lee", 25))
            val output = ByteArrayOutputStream()
            excelOf(users).writeTo(output)
            val bytes = output.toByteArray()

            val headers =
                XSSFWorkbook(ByteArrayInputStream(bytes)).use { workbook ->
                    val row = workbook.getSheetAt(0).getRow(0)
                    (0 until row.lastCellNum).map { row.getCell(it).stringCellValue }
                }
            headers shouldBe listOf("Name", "Age")
            AnnotationExtractor.extractColumns(RoundTripUser::class).map { it.header } shouldBe headers

            parseExcel<RoundTripUser>(ByteArrayInputStream(bytes)).getOrThrow() shouldBe users
        }
    }
})
