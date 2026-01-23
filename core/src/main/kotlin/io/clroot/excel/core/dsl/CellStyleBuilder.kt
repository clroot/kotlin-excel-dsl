@file:Suppress("unused")

package io.clroot.excel.core.dsl

import io.clroot.excel.core.style.Alignment
import io.clroot.excel.core.style.BorderStyle
import io.clroot.excel.core.style.CellStyle
import io.clroot.excel.core.style.Color

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
