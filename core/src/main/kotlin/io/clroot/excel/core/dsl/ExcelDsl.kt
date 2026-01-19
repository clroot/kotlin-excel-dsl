package io.clroot.excel.core.dsl

import io.clroot.excel.core.model.*
import io.clroot.excel.core.style.*

/**
 * DSL marker to prevent scope leakage.
 */
@DslMarker
annotation class ExcelDslMarker

/**
 * Theme interface for DSL.
 * Actual themes are defined in the theme module.
 */
interface ExcelTheme {
    val headerStyle: CellStyle
    val bodyStyle: CellStyle
}

/**
 * Entry point for building an Excel document using DSL.
 */
fun excel(block: ExcelBuilder.() -> Unit): ExcelDocument {
    return ExcelBuilder().apply(block).build()
}

/**
 * Entry point for building an Excel document with a theme.
 */
fun excel(
    theme: ExcelTheme,
    block: ExcelBuilder.() -> Unit,
): ExcelDocument {
    return ExcelBuilder(theme).apply(block).build()
}

/**
 * Builder for constructing an ExcelDocument.
 */
@ExcelDslMarker
class ExcelBuilder(private val theme: ExcelTheme? = null) {
    private val sheets = mutableListOf<Sheet>()
    private var stylesConfig: StylesConfig? = null

    fun <T> sheet(
        name: String,
        block: SheetBuilder<T>.() -> Unit,
    ) {
        sheets.add(SheetBuilder<T>(name).apply(block).build())
    }

    fun styles(block: StylesBuilder.() -> Unit) {
        stylesConfig = StylesBuilder().apply(block).build()
    }

    internal fun build(): ExcelDocument {
        val headerStyle = stylesConfig?.headerStyle ?: theme?.headerStyle
        val bodyStyle = stylesConfig?.bodyStyle ?: theme?.bodyStyle

        // Convert DSL ColumnStyleConfig to model ColumnStyleConfig
        val columnStyles =
            stylesConfig?.columnStyles?.mapValues { (_, config) ->
                io.clroot.excel.core.model.ColumnStyleConfig(
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
 * Builder for constructing a Sheet.
 */
@ExcelDslMarker
class SheetBuilder<T>(private val name: String) {
    private val columns = mutableListOf<ColumnDefinition<T>>()
    private val headerGroups = mutableListOf<HeaderGroup>()
    private val rows = mutableListOf<Row>()

    fun column(
        header: String,
        width: ColumnWidth = ColumnWidth.Auto,
        format: String? = null,
        valueExtractor: (T) -> Any?,
    ) {
        columns.add(ColumnDefinition(header, width, format, null, null, valueExtractor))
    }

    /**
     * Define a column with inline style.
     * Usage: column("금액", style = { bold(); align(RIGHT) }) { it.amount }
     */
    fun column(
        header: String,
        width: ColumnWidth = ColumnWidth.Auto,
        format: String? = null,
        style: CellStyleBuilder.() -> Unit,
        valueExtractor: (T) -> Any?,
    ) {
        val cellStyle = CellStyleBuilder().apply(style).build()
        columns.add(ColumnDefinition(header, width, format, null, cellStyle, valueExtractor))
    }

    /**
     * Define a column with separate header and body styles.
     * Usage: column("금액", headerStyle = { bold() }, bodyStyle = { align(RIGHT) }) { it.amount }
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
        columns.add(ColumnDefinition(header, width, format, hStyle, bStyle, valueExtractor))
    }

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

    fun rows(data: Iterable<T>) {
        data.forEach { item ->
            val cells =
                columns.map { column ->
                    Cell(value = column.valueExtractor(item))
                }
            rows.add(Row(cells))
        }
    }

    internal fun build(): Sheet =
        Sheet(
            name = name,
            columns = columns.toList(),
            headerGroups = headerGroups.toList(),
            rows = rows.toList(),
        )
}

/**
 * Builder for header groups.
 */
@ExcelDslMarker
class HeaderGroupBuilder<T>(private val title: String) {
    private val columns = mutableListOf<ColumnDefinition<T>>()

    fun column(
        header: String,
        width: ColumnWidth = ColumnWidth.Auto,
        format: String? = null,
        valueExtractor: (T) -> Any?,
    ) {
        columns.add(ColumnDefinition(header, width, format, null, null, valueExtractor))
    }

    fun column(
        header: String,
        width: ColumnWidth = ColumnWidth.Auto,
        format: String? = null,
        style: CellStyleBuilder.() -> Unit,
        valueExtractor: (T) -> Any?,
    ) {
        val cellStyle = CellStyleBuilder().apply(style).build()
        columns.add(ColumnDefinition(header, width, format, null, cellStyle, valueExtractor))
    }

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
        columns.add(ColumnDefinition(header, width, format, hStyle, bStyle, valueExtractor))
    }

    internal fun build(): HeaderGroup =
        HeaderGroup(
            title = title,
            columns = columns.toList(),
        )
}

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
