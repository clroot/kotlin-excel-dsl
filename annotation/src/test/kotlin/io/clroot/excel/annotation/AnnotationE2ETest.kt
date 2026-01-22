package io.clroot.excel.annotation

import io.clroot.excel.annotation.style.StyleAlignment
import io.clroot.excel.annotation.style.StyleBorder
import io.clroot.excel.annotation.style.StyleColor
import io.clroot.excel.core.style.CellStyle
import io.clroot.excel.core.style.Color
import io.clroot.excel.render.toByteArray
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.ints.shouldBeGreaterThan
import io.kotest.matchers.shouldBe
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayInputStream

class NegativeStyler : ConditionalStyler<Int> {
    override fun style(value: Int?): CellStyle? =
        if (value != null && value < 0) CellStyle(fontColor = Color.RED) else null
}

class AnnotationE2ETest : DescribeSpec({

    describe("어노테이션 기반 Excel 생성 E2E") {
        @Excel
        @HeaderStyle(bold = true, backgroundColor = StyleColor.BLUE, fontColor = StyleColor.WHITE)
        @BodyStyle(alignment = StyleAlignment.LEFT, border = StyleBorder.THIN)
        data class Product(
            @Column("상품명", order = 1)
            val name: String,

            @Column("가격", format = "#,##0", order = 2)
            @BodyStyle(alignment = StyleAlignment.RIGHT)
            @ConditionalStyle(NegativeStyler::class)
            val price: Int,

            @Column("상태", order = 3)
            @HeaderStyle(backgroundColorHex = "#4CAF50")
            val status: String,
        )

        it("Excel 파일이 정상 생성된다") {
            val products = listOf(
                Product("상품A", 10000, "판매중"),
                Product("상품B", -5000, "할인"),
                Product("상품C", 25000, "품절"),
            )

            val doc = excelOf(products)
            val bytes = doc.toByteArray()

            bytes.size shouldBeGreaterThan 0

            // POI로 검증
            val workbook = XSSFWorkbook(ByteArrayInputStream(bytes))
            val sheet = workbook.getSheetAt(0)

            // 헤더 확인
            val headerRow = sheet.getRow(0)
            headerRow.getCell(0).stringCellValue shouldBe "상품명"
            headerRow.getCell(1).stringCellValue shouldBe "가격"
            headerRow.getCell(2).stringCellValue shouldBe "상태"

            // 데이터 확인
            val dataRow1 = sheet.getRow(1)
            dataRow1.getCell(0).stringCellValue shouldBe "상품A"
            dataRow1.getCell(1).numericCellValue shouldBe 10000.0

            val dataRow2 = sheet.getRow(2)
            dataRow2.getCell(0).stringCellValue shouldBe "상품B"
            dataRow2.getCell(1).numericCellValue shouldBe -5000.0

            workbook.close()
        }

        it("헤더 스타일이 적용된다") {
            val products = listOf(Product("테스트", 100, "테스트"))
            val doc = excelOf(products)
            val bytes = doc.toByteArray()

            val workbook = XSSFWorkbook(ByteArrayInputStream(bytes))
            val sheet = workbook.getSheetAt(0)
            val headerCell = sheet.getRow(0).getCell(0)

            // 헤더 스타일 확인
            val font = workbook.getFontAt(headerCell.cellStyle.fontIndex)
            font.bold shouldBe true

            workbook.close()
        }
    }
})
