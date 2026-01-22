package io.clroot.excel.annotation

import io.clroot.excel.annotation.style.StyleAlignment
import io.clroot.excel.annotation.style.StyleColor
import io.clroot.excel.core.style.Alignment
import io.clroot.excel.core.style.Color
import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.nulls.shouldBeNull
import io.kotest.matchers.nulls.shouldNotBeNull
import io.kotest.matchers.shouldBe

class StyleAnnotationTest : DescribeSpec({

    describe("클래스 레벨 @HeaderStyle") {
        @Excel
        @HeaderStyle(bold = true, backgroundColor = StyleColor.BLUE, fontColor = StyleColor.WHITE)
        data class ClassLevelHeader(
            @Column("이름")
            val name: String,
        )

        it("모든 컬럼 헤더에 스타일이 적용된다") {
            val doc = excelOf(listOf(ClassLevelHeader("테스트")))
            val column = doc.sheets.first().columns.first()

            column.headerStyle.shouldNotBeNull()
            column.headerStyle!!.bold shouldBe true
            column.headerStyle!!.backgroundColor shouldBe Color(0, 0, 255)
            column.headerStyle!!.fontColor shouldBe Color(255, 255, 255)
        }
    }

    describe("클래스 레벨 @BodyStyle") {
        @Excel
        @BodyStyle(alignment = StyleAlignment.CENTER)
        data class ClassLevelBody(
            @Column("값")
            val value: Int,
        )

        it("모든 컬럼 본문에 스타일이 적용된다") {
            val doc = excelOf(listOf(ClassLevelBody(100)))
            val column = doc.sheets.first().columns.first()

            column.bodyStyle.shouldNotBeNull()
            column.bodyStyle!!.alignment shouldBe Alignment.CENTER
        }
    }

    describe("프로퍼티 레벨 스타일 오버라이드") {
        @Excel
        @HeaderStyle(bold = true)
        data class PropertyOverride(
            @Column("이름")
            val name: String,
            @Column("가격")
            @HeaderStyle(backgroundColor = StyleColor.GREEN)
            val price: Int,
        )

        it("프로퍼티 레벨 스타일이 클래스 레벨과 병합된다") {
            val doc = excelOf(listOf(PropertyOverride("상품", 1000)))
            val columns = doc.sheets.first().columns

            // 첫 번째 컬럼: 클래스 레벨만
            columns[0].headerStyle.shouldNotBeNull()
            columns[0].headerStyle!!.bold shouldBe true
            columns[0].headerStyle!!.backgroundColor.shouldBeNull()

            // 두 번째 컬럼: 클래스 + 프로퍼티 병합
            columns[1].headerStyle.shouldNotBeNull()
            columns[1].headerStyle!!.bold shouldBe true // 클래스에서
            columns[1].headerStyle!!.backgroundColor shouldBe Color(0, 128, 0) // 프로퍼티에서
        }
    }

    describe("HEX 색상") {
        @Excel
        @HeaderStyle(backgroundColorHex = "#4A90D9")
        data class HexColor(
            @Column("값")
            val value: String,
        )

        it("HEX 색상이 적용된다") {
            val doc = excelOf(listOf(HexColor("test")))
            val column = doc.sheets.first().columns.first()

            column.headerStyle.shouldNotBeNull()
            column.headerStyle!!.backgroundColor shouldBe Color(74, 144, 217)
        }
    }

    describe("@ConditionalStyle") {
        class PriceStyler : ConditionalStyler<Int> {
            override fun style(value: Int?): io.clroot.excel.core.style.CellStyle? =
                when {
                    value == null -> null
                    value < 0 -> io.clroot.excel.core.style.CellStyle(fontColor = Color.RED)
                    value > 1000 -> io.clroot.excel.core.style.CellStyle(fontColor = Color.GREEN)
                    else -> null
                }
        }

        @Excel
        data class WithConditionalStyle(
            @Column("가격")
            @ConditionalStyle(PriceStyler::class)
            val price: Int,
        )

        @Excel
        data class Transaction(
            @Column("내역", order = 1)
            val description: String,
            @Column("금액", order = 2)
            @ConditionalStyle(PriceStyler::class)
            val amount: Int,
        )

        it("ConditionalStyle이 ColumnDefinition에 설정된다") {
            val doc = excelOf(listOf(WithConditionalStyle(-100)))
            val column = doc.sheets.first().columns.first()

            column.conditionalStyle.shouldNotBeNull()
            // 조건부 스타일이 음수에 대해 빨간색을 반환하는지 확인
            val style = column.conditionalStyle!!.evaluate(-100)
            style.shouldNotBeNull()
            style.fontColor shouldBe Color.RED
        }

        it("음수 값에 빨간색 스타일이 적용된다") {
            val doc = excelOf(listOf(Transaction("테스트", -500)))
            val amountColumn = doc.sheets.first().columns[1]

            amountColumn.conditionalStyle.shouldNotBeNull()
            val style = amountColumn.conditionalStyle!!.evaluate(-500)
            style.shouldNotBeNull()
            style.fontColor shouldBe Color.RED
        }

        it("1000 초과 값에 녹색 스타일이 적용된다") {
            val doc = excelOf(listOf(Transaction("테스트", 2000)))
            val amountColumn = doc.sheets.first().columns[1]

            amountColumn.conditionalStyle.shouldNotBeNull()
            val style = amountColumn.conditionalStyle!!.evaluate(2000)
            style.shouldNotBeNull()
            style.fontColor shouldBe Color.GREEN
        }

        it("조건에 맞지 않으면 null 반환") {
            val doc = excelOf(listOf(Transaction("테스트", 500)))
            val amountColumn = doc.sheets.first().columns[1]

            amountColumn.conditionalStyle.shouldNotBeNull()
            amountColumn.conditionalStyle!!.evaluate(500).shouldBeNull()
        }

        it("null 값도 처리된다") {
            val doc = excelOf(listOf(Transaction("테스트", 100)))
            val amountColumn = doc.sheets.first().columns[1]

            amountColumn.conditionalStyle.shouldNotBeNull()
            amountColumn.conditionalStyle!!.evaluate(null).shouldBeNull()
        }
    }

    describe("테마와 어노테이션 스타일 병합") {
        @Excel
        @HeaderStyle(italic = true)
        data class WithTheme(
            @Column("이름")
            val name: String,
        )

        it("테마 < 클래스 레벨 순서로 병합된다") {
            val theme =
                object : io.clroot.excel.core.dsl.ExcelTheme {
                    override val headerStyle = io.clroot.excel.core.style.CellStyle(bold = true)
                    override val bodyStyle = io.clroot.excel.core.style.CellStyle()
                }
            val doc = excelOf(listOf(WithTheme("테스트")), theme = theme)
            val column = doc.sheets.first().columns.first()

            column.headerStyle.shouldNotBeNull()
            column.headerStyle!!.bold shouldBe true // 테마에서
            column.headerStyle!!.italic shouldBe true // 클래스에서
        }
    }
})
