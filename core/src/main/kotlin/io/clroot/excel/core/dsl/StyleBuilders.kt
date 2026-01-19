@file:Suppress("unused")

package io.clroot.excel.core.dsl

import io.clroot.excel.core.style.Alignment
import io.clroot.excel.core.style.BorderStyle
import io.clroot.excel.core.style.CellStyle
import io.clroot.excel.core.style.Color

/**
 * Configuration for styles.
 */
data class StylesConfig(
    val headerStyle: CellStyle? = null,
    val bodyStyle: CellStyle? = null,
    val columnStyles: Map<String, ColumnStyleConfig> = emptyMap(),
)

/**
 * Style configuration for a specific column.
 */
data class ColumnStyleConfig(
    val headerStyle: CellStyle? = null,
    val bodyStyle: CellStyle? = null,
)

/**
 * Builder for styles DSL.
 */
@ExcelDslMarker
class StylesBuilder {
    private var headerStyle: CellStyle? = null
    private var bodyStyle: CellStyle? = null
    private val columnStyles = mutableMapOf<String, ColumnStyleConfig>()

    fun header(block: CellStyleBuilder.() -> Unit) {
        headerStyle = CellStyleBuilder().apply(block).build()
    }

    fun body(block: CellStyleBuilder.() -> Unit) {
        bodyStyle = CellStyleBuilder().apply(block).build()
    }

    /**
     * Define styles for a specific column by header name.
     * Usage: styles { column("금액") { header { bold() }; body { align(RIGHT) } } }
     */
    fun column(
        headerName: String,
        block: ColumnStyleBuilder.() -> Unit,
    ) {
        columnStyles[headerName] = ColumnStyleBuilder().apply(block).build()
    }

    internal fun build(): StylesConfig = StylesConfig(headerStyle, bodyStyle, columnStyles.toMap())
}

/**
 * Builder for column-specific styles.
 */
@ExcelDslMarker
class ColumnStyleBuilder {
    private var headerStyle: CellStyle? = null
    private var bodyStyle: CellStyle? = null

    fun header(block: CellStyleBuilder.() -> Unit) {
        headerStyle = CellStyleBuilder().apply(block).build()
    }

    fun body(block: CellStyleBuilder.() -> Unit) {
        bodyStyle = CellStyleBuilder().apply(block).build()
    }

    internal fun build(): ColumnStyleConfig = ColumnStyleConfig(headerStyle, bodyStyle)
}

/**
 * Builder for CellStyle DSL.
 */
@ExcelDslMarker
class CellStyleBuilder {
    private var backgroundColor: Color? = null
    private var fontColor: Color? = null
    private var bold: Boolean = false
    private var italic: Boolean = false
    private var alignment: Alignment? = null
    private var border: BorderStyle? = null
    private var numberFormat: String? = null

    fun backgroundColor(color: Color) {
        backgroundColor = color
    }

    fun fontColor(color: Color) {
        fontColor = color
    }

    fun bold() {
        bold = true
    }

    fun italic() {
        italic = true
    }

    fun align(alignment: Alignment) {
        this.alignment = alignment
    }

    fun border(style: BorderStyle) {
        border = style
    }

    fun numberFormat(format: String) {
        numberFormat = format
    }

    internal fun build(): CellStyle =
        CellStyle(
            backgroundColor = backgroundColor,
            fontColor = fontColor,
            bold = bold,
            italic = italic,
            alignment = alignment,
            border = border,
            numberFormat = numberFormat,
        )
}
