@file:Suppress("unused")

package io.clroot.excel.core.dsl

import io.clroot.excel.core.model.ColumnDefinition
import io.clroot.excel.core.model.ColumnStyleConfig
import io.clroot.excel.core.model.ColumnWidth
import io.clroot.excel.core.model.ExcelDocument
import io.clroot.excel.core.model.FreezePane
import io.clroot.excel.core.model.HeaderGroup
import io.clroot.excel.core.model.Sheet

/**
 * Builder for constructing an [ExcelDocument].
 *
 * Use this builder via the [excel] DSL function to create Excel documents
 * with multiple sheets, styles, and data.
 *
 * Example:
 * ```kotlin
 * excel {
 *     sheet<User>("Users") {
 *         column("Name") { it.name }
 *         column("Age") { it.age }
 *         rows(users)
 *     }
 *     styles {
 *         header { bold(); backgroundColor(Color.LIGHT_BLUE) }
 *     }
 * }
 * ```
 *
 * @see excel
 * @see SheetBuilder
 */
@ExcelDslMarker
class ExcelBuilder(private val theme: ExcelTheme? = null) {
    private val sheets = mutableListOf<Sheet>()
    private var stylesConfig: StylesConfig? = null

    /**
     * Adds a new sheet to the Excel document.
     *
     * @param T the type of data objects that will populate this sheet
     * @param name the name of the sheet (displayed as tab name in Excel)
     * @param block the builder block to configure columns and rows
     */
    fun <T> sheet(
        name: String,
        block: SheetBuilder<T>.() -> Unit,
    ) {
        sheets.add(SheetBuilder<T>(name).apply(block).build())
    }

    /**
     * Configures global styles for the Excel document.
     *
     * Styles defined here apply to all sheets unless overridden at the column level.
     *
     * @param block the builder block to configure header, body, and column-specific styles
     * @see StylesBuilder
     */
    fun styles(block: StylesBuilder.() -> Unit) {
        stylesConfig = StylesBuilder().apply(block).build()
    }

    internal fun build(): ExcelDocument {
        val headerStyle = stylesConfig?.headerStyle ?: theme?.headerStyle
        val bodyStyle = stylesConfig?.bodyStyle ?: theme?.bodyStyle

        // Convert DSL ColumnStyleConfig to model ColumnStyleConfig
        val columnStyles =
            stylesConfig?.columnStyles?.mapValues { (_, config) ->
                ColumnStyleConfig(
                    headerStyle = config.headerStyle,
                    bodyStyle = config.bodyStyle,
                )
            } ?: emptyMap()

        return ExcelDocument(
            sheets = sheets.toList(),
            headerStyle = headerStyle,
            bodyStyle = bodyStyle,
            columnStyles = columnStyles,
        )
    }
}

/**
 * Builder for constructing a [Sheet] within an Excel document.
 *
 * Use this builder to define columns, header groups, and populate data rows.
 *
 * Example:
 * ```kotlin
 * sheet<Product>("Products") {
 *     column("Name", width = 30.chars) { it.name }
 *     column("Price", format = "#,##0") { it.price }
 *     headerGroup("Details") {
 *         column("Category") { it.category }
 *         column("Stock") { it.stock }
 *     }
 *     rows(products)
 * }
 * ```
 *
 * @param T the type of data objects that will populate this sheet
 * @see ColumnDefinition
 * @see HeaderGroup
 */
@ExcelDslMarker
class SheetBuilder<T>(private val name: String) {
    private val columns = mutableListOf<ColumnDefinition<T>>()
    private val headerGroups = mutableListOf<HeaderGroup>()
    private var dataSource: Iterable<T>? = null
    private var freezePaneConfig: FreezePane? = null
    private var autoFilterEnabled: Boolean = false
    private var alternateStyle: io.clroot.excel.core.style.CellStyle? = null

    /**
     * Freezes rows and/or columns in the sheet.
     *
     * Frozen panes remain visible while scrolling through the rest of the sheet.
     * This is useful for keeping headers visible when viewing large datasets.
     *
     * Example:
     * ```kotlin
     * sheet<User>("Users") {
     *     freezePane(row = 1)  // Freeze header row
     *     freezePane(row = 1, col = 1)  // Freeze header row and first column
     * }
     * ```
     *
     * @param row the number of rows to freeze from the top (default: 0, must be non-negative)
     * @param col the number of columns to freeze from the left (default: 0, must be non-negative)
     * @throws IllegalArgumentException if row or col is negative
     */
    fun freezePane(
        row: Int = 0,
        col: Int = 0,
    ) {
        require(row >= 0) { "freezePane row must be non-negative, but was $row" }
        require(col >= 0) { "freezePane col must be non-negative, but was $col" }
        freezePaneConfig = FreezePane(row, col)
    }

    /**
     * Enables auto-filter dropdown on the header row.
     *
     * When enabled, Excel will show filter dropdowns on each column header,
     * allowing users to filter and sort data.
     *
     * Example:
     * ```kotlin
     * sheet<User>("Users") {
     *     autoFilter()
     *     column("Name") { it.name }
     *     column("Age") { it.age }
     * }
     * ```
     */
    fun autoFilter() {
        autoFilterEnabled = true
    }

    /**
     * Applies alternating row styling (zebra stripes) to even-numbered data rows.
     *
     * This creates a visual pattern that makes it easier to read across rows.
     * The style is applied to rows at indices 0, 2, 4, etc. (0-indexed from data rows).
     *
     * Example:
     * ```kotlin
     * sheet<User>("Users") {
     *     alternateRowStyle {
     *         backgroundColor(Color.LIGHT_GRAY)
     *     }
     *     column("Name") { it.name }
     *     rows(users)
     * }
     * ```
     *
     * @param block the style builder block for alternate rows
     */
    fun alternateRowStyle(block: CellStyleBuilder.() -> Unit) {
        alternateStyle = CellStyleBuilder().apply(block).build()
    }

    /**
     * Defines a column in the sheet.
     *
     * @param header the column header text displayed in the first row
     * @param width the column width specification (default: [ColumnWidth.Auto])
     * @param format the number format pattern (e.g., "#,##0", "yyyy-MM-dd")
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
     * Example:
     * ```kotlin
     * column("Price", style = { bold(); align(RIGHT) }) { it.price }
     * ```
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
     * Example:
     * ```kotlin
     * column(
     *     "Amount",
     *     headerStyle = { bold(); backgroundColor(Color.LIGHT_BLUE) },
     *     bodyStyle = { align(RIGHT); numberFormat("#,##0") }
     * ) { it.amount }
     * ```
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

    /**
     * Creates a header group with a spanning title row.
     *
     * Header groups create a merged cell above the grouped columns,
     * useful for categorizing related columns.
     *
     * Example:
     * ```kotlin
     * headerGroup("Contact Info") {
     *     column("Email") { it.email }
     *     column("Phone") { it.phone }
     * }
     * ```
     *
     * @param title the title text for the header group (displayed in merged cell)
     * @param block the builder block to define columns within this group
     */
    fun headerGroup(
        title: String,
        block: HeaderGroupBuilder<T>.() -> Unit,
    ) {
        val builder = HeaderGroupBuilder<T>(title)
        builder.apply(block)
        val group = builder.build()
        headerGroups.add(group)
        // Also add columns to the flat list for data extraction
        @Suppress("UNCHECKED_CAST")
        columns.addAll(group.columns.map { it as ColumnDefinition<T> })
    }

    /**
     * Populates the sheet with data rows.
     *
     * The data is stored as-is and processed during rendering (streaming mode).
     * This enables memory-efficient handling of large datasets.
     *
     * @param data the iterable of data objects to populate as rows
     */
    fun rows(data: Iterable<T>) {
        this.dataSource = data
    }

    internal fun build(): Sheet =
        Sheet(
            name = name,
            columns = columns.toList(),
            headerGroups = headerGroups.toList(),
            dataSource = dataSource,
            freezePane = freezePaneConfig,
            autoFilter = autoFilterEnabled,
            alternateRowStyle = alternateStyle,
        )
}

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
