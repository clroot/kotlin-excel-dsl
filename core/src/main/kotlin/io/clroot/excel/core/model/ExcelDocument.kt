@file:Suppress("unused")

package io.clroot.excel.core.model

import io.clroot.excel.core.style.CellStyle

/**
 * Represents an Excel document containing multiple sheets.
 */
data class ExcelDocument(
    val sheets: List<Sheet> = emptyList(),
    val headerStyle: CellStyle? = null,
    val bodyStyle: CellStyle? = null,
    val columnStyles: Map<String, ColumnStyleConfig> = emptyMap(),
)

/**
 * Style configuration for a specific column (header name -> styles).
 */
data class ColumnStyleConfig(
    val headerStyle: CellStyle? = null,
    val bodyStyle: CellStyle? = null,
)

/**
 * Configuration for freezing panes (rows and columns).
 *
 * @property row the number of rows to freeze from the top (0 = no frozen rows)
 * @property col the number of columns to freeze from the left (0 = no frozen columns)
 */
data class FreezePane(
    val row: Int = 0,
    val col: Int = 0,
)

/**
 * Represents a single sheet in an Excel document.
 *
 * Uses streaming mode with [dataSource] and column extractors for memory-efficient large dataset handling.
 *
 * @property freezePane configuration for frozen rows/columns (null = no freezing)
 * @property autoFilter whether to enable auto-filter on the header row
 * @property alternateRowStyle style applied to even-numbered data rows (0-indexed: rows 0, 2, 4...)
 */
data class Sheet(
    val name: String,
    val columns: List<ColumnDefinition<*>> = emptyList(),
    val headerGroups: List<HeaderGroup> = emptyList(),
    val dataSource: Iterable<*>? = null,
    val freezePane: FreezePane? = null,
    val autoFilter: Boolean = false,
    val alternateRowStyle: CellStyle? = null,
)

/**
 * Represents a header group for multi-row headers.
 */
data class HeaderGroup(
    val title: String,
    val columns: List<ColumnDefinition<*>>,
)

/**
 * Defines a column with header, width, and value extraction logic.
 */
data class ColumnDefinition<T>(
    val header: String,
    val width: ColumnWidth = ColumnWidth.Auto,
    val format: String? = null,
    val headerStyle: io.clroot.excel.core.style.CellStyle? = null,
    val bodyStyle: io.clroot.excel.core.style.CellStyle? = null,
    val valueExtractor: (T) -> Any?,
)

/**
 * Represents column width configuration.
 */
sealed class ColumnWidth {
    data object Auto : ColumnWidth()

    data class Fixed(val chars: Int) : ColumnWidth()
}

/**
 * Extension property to create Fixed column width.
 * Usage: column("이름", width = 20.chars) { it.name }
 */
val Int.chars: ColumnWidth get() = ColumnWidth.Fixed(this)

/**
 * Auto column width constant.
 */
val auto: ColumnWidth = ColumnWidth.Auto
