package io.clroot.excel.render

import io.clroot.excel.core.model.ColumnDefinition
import io.clroot.excel.core.model.ColumnWidth

/**
 * Tracks maximum column widths for auto-width calculation.
 *
 * This class maintains O(1) memory usage by only tracking the maximum width
 * encountered for each auto-width column, rather than storing all cell values.
 *
 * @property columns the list of column definitions
 */
internal class ColumnWidthTracker(
    private val columns: List<ColumnDefinition<*>>,
) {
    private val autoWidthColumnIndices: Set<Int> =
        columns.mapIndexedNotNull { index, column ->
            if (column.width is ColumnWidth.Auto) index else null
        }.toSet()

    private val maxWidths: MutableMap<Int, Double> =
        autoWidthColumnIndices.associateWith { index ->
            ColumnWidthCalculator.calculateTextWidth(columns[index].header)
        }.toMutableMap()

    /**
     * Updates the tracked width for a column if the new value is wider.
     *
     * @param columnIndex the index of the column
     * @param value the cell value (will be converted to string for width calculation)
     */
    fun trackWidth(
        columnIndex: Int,
        value: Any?,
    ) {
        if (columnIndex !in autoWidthColumnIndices) return

        val textWidth = ColumnWidthCalculator.calculateTextWidth(value?.toString() ?: "")
        val currentMax = maxWidths[columnIndex] ?: 0.0
        if (textWidth > currentMax) {
            maxWidths[columnIndex] = textWidth
        }
    }

    /**
     * Gets the calculated width for a column in POI units.
     *
     * @param columnIndex the index of the column
     * @return the width in POI units (1/256 of a character), or null if not an auto-width column
     */
    fun getPoiWidth(columnIndex: Int): Int? {
        val maxWidth = maxWidths[columnIndex] ?: return null
        return ColumnWidthCalculator.toPoiWidth(maxWidth)
    }
}
