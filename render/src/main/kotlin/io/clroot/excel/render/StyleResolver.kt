package io.clroot.excel.render

import io.clroot.excel.core.model.ColumnDefinition
import io.clroot.excel.core.model.ColumnStyleConfig
import io.clroot.excel.core.model.ExcelDocument
import io.clroot.excel.core.model.Sheet
import io.clroot.excel.core.style.CellStyle

/**
 * Resolves cell styles based on priority rules.
 *
 * Style priority (highest to lowest):
 * 1. Inline styles - defined directly on the column
 * 2. Column-specific styles - defined via `styles { column("Name") { ... } }`
 * 3. Global styles - defined via `styles { header { ... } }` or `styles { body { ... } }`
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
    /**
     * Creates a StyleResolver from an ExcelDocument.
     */
    companion object {
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
     * Resolves the header style for a column.
     *
     * @param column the column definition
     * @return the resolved header style, or null if no style is defined
     */
    fun resolveHeaderStyle(column: ColumnDefinition<*>): CellStyle? {
        return column.headerStyle
            ?: columnStyles[column.header]?.headerStyle
            ?: globalHeaderStyle
    }

    /**
     * Resolves the body style for a column.
     *
     * @param column the column definition
     * @return the resolved body style, or null if no style is defined
     */
    fun resolveBodyStyle(column: ColumnDefinition<*>): CellStyle? {
        return column.bodyStyle
            ?: columnStyles[column.header]?.bodyStyle
            ?: globalBodyStyle
    }

    /**
     * Resolves the final body style for a cell, considering alternate row styling.
     *
     * @param column the column definition
     * @param sheet the sheet containing alternate row style configuration
     * @param isAlternateRow whether this is an alternate (even-indexed) row
     * @return the resolved and merged style, or null if no style is defined
     */
    fun resolveBodyStyleWithAlternate(
        column: ColumnDefinition<*>,
        sheet: Sheet,
        isAlternateRow: Boolean,
    ): CellStyle? {
        val bodyStyle = resolveBodyStyle(column)
        val alternateStyle = sheet.alternateRowStyle

        return if (isAlternateRow && alternateStyle != null) {
            bodyStyle?.merge(alternateStyle) ?: alternateStyle
        } else {
            bodyStyle
        }
    }

    /**
     * Resolves the final style for a cell, including date format and conditional style if applicable.
     *
     * Style priority (lowest to highest):
     * 1. alternateRowStyle
     * 2. globalBodyStyle
     * 3. columnStyle (from styles DSL)
     * 4. inlineBodyStyle
     * 5. conditionalStyle (highest priority)
     *
     * @param column the column definition
     * @param sheet the sheet containing alternate row style configuration
     * @param isAlternateRow whether this is an alternate (even-indexed) row
     * @param dateFormat the date format to apply, or null if not a date cell
     * @param cellValue the cell value to evaluate conditional style against
     * @return the resolved style with date format applied, or null if no style is needed
     */
    fun resolveFinalStyle(
        column: ColumnDefinition<*>,
        sheet: Sheet,
        isAlternateRow: Boolean,
        dateFormat: String?,
        cellValue: Any? = null,
    ): CellStyle? {
        val baseStyle = resolveBodyStyleWithAlternate(column, sheet, isAlternateRow)

        // Evaluate conditional style if present
        val conditionalStyle = column.conditionalStyle?.evaluate(cellValue)

        // Merge base style with conditional style (conditional has highest priority)
        val mergedStyle =
            if (conditionalStyle != null) {
                baseStyle?.merge(conditionalStyle) ?: conditionalStyle
            } else {
                baseStyle
            }

        return if (dateFormat != null) {
            (mergedStyle ?: CellStyle()).copy(numberFormat = dateFormat)
        } else {
            mergedStyle
        }
    }
}
