package io.clroot.excel.annotation

import io.clroot.excel.render.toByteArray
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayInputStream

@Excel
data class DeclarationOrderUser(
    @Column("Name") val name: String,
    @Column("Age") val age: Int,
    @Column("City") val city: String,
)

@Excel
data class MixedOrderUser(
    @Column("Zip") val zip: String,
    @Column("Email", order = 2) val email: String,
    @Column("Name", order = 1) val name: String,
    @Column("Age") val age: Int,
)

class ColumnOrderTest : DescribeSpec({

    fun headersOf(bytes: ByteArray): List<String> =
        XSSFWorkbook(ByteArrayInputStream(bytes)).use { workbook ->
            val row = workbook.getSheetAt(0).getRow(0)
            (0 until row.lastCellNum).map { row.getCell(it).stringCellValue }
        }

    describe("excelOf column order") {
        it("follows primary constructor declaration order when order is omitted") {
            val bytes = excelOf(listOf(DeclarationOrderUser("Kim", 30, "Seoul"))).toByteArray()

            headersOf(bytes) shouldBe listOf("Name", "Age", "City")
        }

        it("places explicit order first, then constructor order for ties") {
            val bytes = excelOf(listOf(MixedOrderUser("12345", "a@b.c", "Kim", 30))).toByteArray()

            headersOf(bytes) shouldBe listOf("Name", "Email", "Zip", "Age")
        }

        it("columnProperties returns the shared ordering") {
            columnProperties(MixedOrderUser::class).map { it.name } shouldBe listOf("name", "email", "zip", "age")
        }
    }
})
