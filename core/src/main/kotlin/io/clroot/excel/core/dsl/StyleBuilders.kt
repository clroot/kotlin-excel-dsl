@file:Suppress("unused")

package io.clroot.excel.core.dsl

import io.clroot.excel.core.style.Alignment
import io.clroot.excel.core.style.BorderStyle
import io.clroot.excel.core.style.CellStyle
import io.clroot.excel.core.style.Color

/**
 * Configuration holding all style definitions for an Excel document.
 *
 * @property headerStyle the default style for all header cells
 * @property bodyStyle the default style for all body cells
 * @property columnStyles column-specific style overrides, keyed by header name
 */
data class StylesConfig(
    val headerStyle: CellStyle? = null,
    val bodyStyle: CellStyle? = null,
    val columnStyles: Map<String, ColumnStyleConfig> = emptyMap(),
)

/**
 * Style configuration for a specific column.
 *
 * @property headerStyle the style for this column's header cell (overrides global header style)
 * @property bodyStyle the style for this column's body cells (overrides global body style)
 */
data class ColumnStyleConfig(
    val headerStyle: CellStyle? = null,
    val bodyStyle: CellStyle? = null,
)

/**
 * Builder for configuring Excel document styles.
 *
 * Use this builder to define global header/body styles and column-specific style overrides.
 *
 * Example:
 * ```kotlin
 * styles {
 *     header {
 *         bold()
 *         backgroundColor(Color.LIGHT_BLUE)
 *     }
 *     body {
 *         fontColor(Color.DARK_GRAY)
 *     }
 *     column("Price") {
 *         body { align(RIGHT); numberFormat("#,##0") }
 *     }
 * }
 * ```
 *
 * @see CellStyleBuilder
 * @see ColumnStyleBuilder
 */
@ExcelDslMarker
class StylesBuilder {
    private var headerStyle: CellStyle? = null
    private var bodyStyle: CellStyle? = null
    private val columnStyles = mutableMapOf<String, ColumnStyleConfig>()

    /**
     * Defines the default style for all header cells.
     *
     * @param block the style builder block
     */
    fun header(block: CellStyleBuilder.() -> Unit) {
        headerStyle = CellStyleBuilder().apply(block).build()
    }

    /**
     * Defines the default style for all body cells.
     *
     * @param block the style builder block
     */
    fun body(block: CellStyleBuilder.() -> Unit) {
        bodyStyle = CellStyleBuilder().apply(block).build()
    }

    /**
     * Defines styles for a specific column by header name.
     *
     * Column-specific styles override the global header/body styles.
     *
     * Example:
     * ```kotlin
     * column("Amount") {
     *     header { bold(); backgroundColor(Color.YELLOW) }
     *     body { align(RIGHT); numberFormat("#,##0") }
     * }
     * ```
     *
     * @param headerName the column header name to apply styles to
     * @param block the column style builder block
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
 *
 * Allows defining separate header and body styles for a specific column.
 *
 * @see StylesBuilder.column
 */
@ExcelDslMarker
class ColumnStyleBuilder {
    private var headerStyle: CellStyle? = null
    private var bodyStyle: CellStyle? = null

    /**
     * Defines the header cell style for this column.
     *
     * @param block the style builder block
     */
    fun header(block: CellStyleBuilder.() -> Unit) {
        headerStyle = CellStyleBuilder().apply(block).build()
    }

    /**
     * Defines the body cell style for this column.
     *
     * @param block the style builder block
     */
    fun body(block: CellStyleBuilder.() -> Unit) {
        bodyStyle = CellStyleBuilder().apply(block).build()
    }

    internal fun build(): ColumnStyleConfig = ColumnStyleConfig(headerStyle, bodyStyle)
}

/**
 * Builder for constructing a [CellStyle].
 *
 * Provides a fluent DSL for defining cell appearance including
 * colors, fonts, alignment, borders, and number formats.
 *
 * Example:
 * ```kotlin
 * CellStyleBuilder().apply {
 *     backgroundColor(Color.LIGHT_BLUE)
 *     fontColor(Color.DARK_GRAY)
 *     bold()
 *     align(Alignment.CENTER)
 *     border(BorderStyle.THIN)
 *     numberFormat("#,##0.00")
 * }.build()
 * ```
 *
 * @see CellStyle
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

    /**
     * Sets the cell background color.
     *
     * @param color the background color
     */
    fun backgroundColor(color: Color) {
        backgroundColor = color
    }

    /**
     * Sets the font color.
     *
     * @param color the font color
     */
    fun fontColor(color: Color) {
        fontColor = color
    }

    /**
     * Makes the font bold.
     */
    fun bold() {
        bold = true
    }

    /**
     * Makes the font italic.
     */
    fun italic() {
        italic = true
    }

    /**
     * Sets the horizontal text alignment.
     *
     * @param alignment the alignment (LEFT, CENTER, RIGHT)
     */
    fun align(alignment: Alignment) {
        this.alignment = alignment
    }

    /**
     * Sets the cell border style.
     *
     * @param style the border style (NONE, THIN, MEDIUM, THICK, DOUBLE)
     */
    fun border(style: BorderStyle) {
        border = style
    }

    /**
     * Sets the number format pattern.
     *
     * Common patterns:
     * - `#,##0` - Integer with thousands separator
     * - `#,##0.00` - Decimal with 2 places
     * - `yyyy-MM-dd` - Date format
     * - `0%` - Percentage
     *
     * @param format the Excel number format pattern
     */
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
