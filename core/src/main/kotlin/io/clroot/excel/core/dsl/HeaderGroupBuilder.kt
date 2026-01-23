@file:Suppress("unused")

package io.clroot.excel.core.dsl

import io.clroot.excel.core.model.ColumnDefinition
import io.clroot.excel.core.model.ColumnWidth
import io.clroot.excel.core.model.HeaderGroup

/**
 * Builder for creating header groups within a sheet.
 *
 * Header groups allow grouping related columns under a common title
 * that spans across multiple columns in the header row.
 *
 * @param T the type of data objects for value extraction
 * @see SheetBuilder.headerGroup
 */
@ExcelDslMarker
class HeaderGroupBuilder<T>(private val title: String) {
    private val columns = mutableListOf<ColumnDefinition<T>>()

    /**
     * Defines a column within the header group.
     *
     * @param header the column header text
     * @param width the column width specification
     * @param format the number format pattern
     * @param valueExtractor a function to extract the cell value from each data object
     */
    fun column(
        header: String,
        width: ColumnWidth = ColumnWidth.Auto,
        format: String? = null,
        valueExtractor: (T) -> Any?,
    ) {
        columns.add(
            ColumnDefinition(
                header = header,
                width = width,
                format = format,
                headerStyle = null,
                bodyStyle = null,
                conditionalStyle = null,
                valueExtractor = valueExtractor,
            ),
        )
    }

    /**
     * Defines a column with inline style applied to body cells.
     *
     * @param header the column header text
     * @param width the column width specification
     * @param format the number format pattern
     * @param style the style builder block for body cells
     * @param valueExtractor a function to extract the cell value from each data object
     */
    fun column(
        header: String,
        width: ColumnWidth = ColumnWidth.Auto,
        format: String? = null,
        style: CellStyleBuilder.() -> Unit,
        valueExtractor: (T) -> Any?,
    ) {
        val cellStyle = CellStyleBuilder().apply(style).build()
        columns.add(
            ColumnDefinition(
                header = header,
                width = width,
                format = format,
                headerStyle = null,
                bodyStyle = cellStyle,
                conditionalStyle = null,
                valueExtractor = valueExtractor,
            ),
        )
    }

    /**
     * Defines a column with separate header and body styles.
     *
     * @param header the column header text
     * @param width the column width specification
     * @param format the number format pattern
     * @param headerStyle the style builder block for the header cell (optional)
     * @param bodyStyle the style builder block for body cells (optional)
     * @param valueExtractor a function to extract the cell value from each data object
     */
    fun column(
        header: String,
        width: ColumnWidth = ColumnWidth.Auto,
        format: String? = null,
        headerStyle: (CellStyleBuilder.() -> Unit)? = null,
        bodyStyle: (CellStyleBuilder.() -> Unit)? = null,
        valueExtractor: (T) -> Any?,
    ) {
        val hStyle = headerStyle?.let { CellStyleBuilder().apply(it).build() }
        val bStyle = bodyStyle?.let { CellStyleBuilder().apply(it).build() }
        columns.add(
            ColumnDefinition(
                header = header,
                width = width,
                format = format,
                headerStyle = hStyle,
                bodyStyle = bStyle,
                conditionalStyle = null,
                valueExtractor = valueExtractor,
            ),
        )
    }

    internal fun build(): HeaderGroup =
        HeaderGroup(
            title = title,
            columns = columns.toList(),
        )
}
