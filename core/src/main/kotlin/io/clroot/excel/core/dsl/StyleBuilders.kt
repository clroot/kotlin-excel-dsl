@file:Suppress("unused")

package io.clroot.excel.core.dsl

import io.clroot.excel.core.style.CellStyle

/**
 * Configuration holding all style definitions for an Excel document.
 *
 * @property headerStyle the default style for all header cells
 * @property bodyStyle the default style for all body cells
 * @property columnStyles column-specific styles merged over the global styles, keyed by header name
 */
data class StylesConfig(
    val headerStyle: CellStyle? = null,
    val bodyStyle: CellStyle? = null,
    val columnStyles: Map<String, ColumnStyleConfig> = emptyMap(),
)

/**
 * Style configuration for a specific column.
 *
 * @property headerStyle the style for this column's header cell (merged over the global header style)
 * @property bodyStyle the style for this column's body cells (merged over the global body style)
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
     * Column-specific styles are merged over the global header/body styles: properties set here
     * win, and properties left unset are inherited from the global styles.
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
