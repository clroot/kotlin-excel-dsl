package io.clroot.excel.render

import io.clroot.excel.core.model.ColumnDefinition
import io.clroot.excel.core.model.ColumnStyleConfig
import io.clroot.excel.core.model.ExcelDocument
import io.clroot.excel.core.style.CellStyle
import java.time.LocalDate
import java.time.LocalDateTime

/**
 * Resolves cell styles by merging style layers.
 *
 * Layers are merged with [CellStyle.merge]: a property set (non-null) in a higher layer
 * overrides the same property of lower layers, and properties a layer leaves unset are
 * inherited. No layer replaces another as a whole.
 *
 * Header layers (lowest to highest):
 * 1. Global header style - theme header merged with `styles { header { ... } }`
 * 2. Column-specific header style - `styles { column("Name") { header { ... } } }`
 * 3. Inline header style - `column(..., headerStyle = { ... })`
 *
 * Body layers (lowest to highest):
 * 1. Alternate row style - `alternateRowStyle { ... }`, even-indexed data rows only
 * 2. Global body style - theme body merged with `styles { body { ... } }`
 * 3. Column-specific body style - `styles { column("Name") { body { ... } } }`
 * 4. Inline body style - `column(..., style = { ... })` or `bodyStyle = { ... }`
 * 5. Column format - `column(..., format = "...")` / `@Column(format = "...")` as numberFormat
 * 6. Conditional style - evaluated per cell value
 *
 * `LocalDate` and `LocalDateTime` cells get [DEFAULT_DATE_FORMAT] / [DEFAULT_DATE_TIME_FORMAT]
 * only when no layer sets a numberFormat.
 *
 * @property globalHeaderStyle global header style from document
 * @property globalBodyStyle global body style from document
 * @property columnStyles column-specific styles from document
 */
internal class StyleResolver(
    private val globalHeaderStyle: CellStyle?,
    private val globalBodyStyle: CellStyle?,
    private val columnStyles: Map<String, ColumnStyleConfig>,
) {
    companion object {
        /** Number format applied to [LocalDate] cells without a user number format. */
        const val DEFAULT_DATE_FORMAT = "yyyy-mm-dd"

        /** Number format applied to [LocalDateTime] cells without a user number format. */
        const val DEFAULT_DATE_TIME_FORMAT = "yyyy-mm-dd hh:mm:ss"

        /**
         * Creates a StyleResolver from an ExcelDocument.
         */
        fun from(document: ExcelDocument): StyleResolver =
            StyleResolver(
                globalHeaderStyle = document.headerStyle,
                globalBodyStyle = document.bodyStyle,
                columnStyles = document.columnStyles,
            )
    }

    /**
     * Resolves the style for group headers.
     *
     * Group headers are not tied to any specific column, so they always use
     * the global header style regardless of column-specific or inline styles.
     *
     * @return the global header style, or null if not defined
     */
    fun resolveGroupHeaderStyle(): CellStyle? = globalHeaderStyle

    /**
     * Resolves the header style for a column by merging global < column-specific < inline.
     *
     * @param column the column definition
     * @return the resolved header style, or null if no style is defined
     */
    fun resolveHeaderStyle(column: ColumnDefinition<*>): CellStyle? =
        mergeLayers(
            globalHeaderStyle,
            columnStyles[column.header]?.headerStyle,
            column.headerStyle,
        )

    /**
     * Resolves the static part of a column's body style: every layer except the conditional style.
     *
     * The result does not depend on the cell value, so callers can compute it once per column.
     *
     * @param column the column definition
     * @param alternateRowStyle the alternate row style when rendering an alternate row, otherwise null
     * @return the resolved body style, or null if no style is defined
     */
    fun resolveBodyStyle(
        column: ColumnDefinition<*>,
        alternateRowStyle: CellStyle? = null,
    ): CellStyle? =
        mergeLayers(
            alternateRowStyle,
            globalBodyStyle,
            columnStyles[column.header]?.bodyStyle,
            column.bodyStyle,
            column.format?.let { CellStyle(numberFormat = it) },
        )

    /**
     * Resolves the final style for a body cell.
     *
     * Merges the conditional style (highest priority) on top of [baseStyle] and, for
     * `LocalDate`/`LocalDateTime` values, applies the default date format when no layer
     * set a numberFormat. See the class documentation for the full layer order.
     *
     * @param column the column definition
     * @param baseStyle the result of [resolveBodyStyle] for this column and row kind
     * @param cellValue the cell value to evaluate the conditional style against
     * @return the resolved style, or null if no style is needed
     */
    fun resolveFinalStyle(
        column: ColumnDefinition<*>,
        baseStyle: CellStyle?,
        cellValue: Any?,
    ): CellStyle? {
        val conditionalStyle = column.conditionalStyle?.evaluate(cellValue)
        val mergedStyle =
            if (conditionalStyle != null) baseStyle?.merge(conditionalStyle) ?: conditionalStyle else baseStyle

        val defaultFormat = defaultNumberFormat(cellValue) ?: return mergedStyle
        return when {
            mergedStyle == null -> CellStyle(numberFormat = defaultFormat)
            mergedStyle.numberFormat == null -> mergedStyle.copy(numberFormat = defaultFormat)
            else -> mergedStyle
        }
    }

    private fun defaultNumberFormat(value: Any?): String? =
        when (value) {
            is LocalDate -> DEFAULT_DATE_FORMAT
            is LocalDateTime -> DEFAULT_DATE_TIME_FORMAT
            else -> null
        }

    private fun mergeLayers(vararg layers: CellStyle?): CellStyle? {
        var result: CellStyle? = null
        for (layer in layers) {
            if (layer == null) continue
            result = result?.merge(layer) ?: layer
        }
        return result
    }
}
