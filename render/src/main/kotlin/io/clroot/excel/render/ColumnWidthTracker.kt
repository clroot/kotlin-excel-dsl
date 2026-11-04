package io.clroot.excel.render

import io.clroot.excel.core.model.ColumnDefinition
import io.clroot.excel.core.model.ColumnWidth
import io.clroot.excel.core.model.Formula
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.DateUtil
import java.math.BigDecimal
import java.math.BigInteger
import java.time.LocalDate
import java.time.LocalDateTime
import java.util.Locale
import kotlin.math.abs

/**
 * Tracks maximum column widths for auto-width calculation.
 *
 * Widths are estimated from the text Excel displays for each cell rather than from the raw
 * value's `toString()`:
 * - formulas use a fixed estimate ([FORMULA_WIDTH_CHARS]) because their result is unknown
 * - numbers use plain notation (no exponent), or the applied number format when one is set
 * - dates use the applied date format
 *
 * Memory usage is O(number of columns): only the maximum width of each auto-width column is
 * kept, never the cell values.
 *
 * @property columns the list of column definitions
 */
internal class ColumnWidthTracker(
    private val columns: List<ColumnDefinition<*>>,
) {
    companion object {
        /** Estimated display width, in characters, of a formula cell's result. */
        const val FORMULA_WIDTH_CHARS = 10

        private const val GENERAL_FORMAT = "General"

        /** Integral doubles below this magnitude are displayed without a fraction or exponent. */
        private const val MAX_EXACT_INTEGRAL_DOUBLE = 1e15
    }

    private val autoWidthColumnIndices: Set<Int> =
        columns.mapIndexedNotNull { index, column ->
            if (column.width is ColumnWidth.Auto) index else null
        }.toSet()

    private val maxWidths: MutableMap<Int, Double> =
        autoWidthColumnIndices.associateWith { index ->
            ColumnWidthCalculator.calculateTextWidth(columns[index].header)
        }.toMutableMap()

    // Created on first use; caches parsed formats per format string (O(distinct formats))
    private val formatter by lazy(LazyThreadSafetyMode.NONE) { DataFormatter(Locale.US) }

    /**
     * Updates the tracked width for a column if the new value is wider.
     *
     * @param columnIndex the index of the column
     * @param value the cell value
     * @param numberFormat the number format applied to the cell, or null if none
     */
    fun trackWidth(
        columnIndex: Int,
        value: Any?,
        numberFormat: String? = null,
    ) {
        if (columnIndex !in autoWidthColumnIndices) return

        val textWidth =
            if (value is Formula) {
                FORMULA_WIDTH_CHARS.toDouble()
            } else {
                ColumnWidthCalculator.calculateTextWidth(displayText(value, numberFormat))
            }
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

    private fun displayText(
        value: Any?,
        numberFormat: String?,
    ): String =
        when (value) {
            null -> ""
            is String -> value
            is Boolean -> if (value) "TRUE" else "FALSE"
            is Number -> numberText(value, numberFormat)
            is LocalDate ->
                dateText(numberFormat ?: StyleResolver.DEFAULT_DATE_FORMAT) { DateUtil.getExcelDate(value) }
            is LocalDateTime ->
                dateText(numberFormat ?: StyleResolver.DEFAULT_DATE_TIME_FORMAT) { DateUtil.getExcelDate(value) }
            else -> value.toString()
        }

    private fun numberText(
        value: Number,
        numberFormat: String?,
    ): String =
        if (numberFormat == null || numberFormat.equals(GENERAL_FORMAT, ignoreCase = true)) {
            plainText(value)
        } else {
            formatter.formatRawCellContents(value.toDouble(), -1, numberFormat)
        }

    private inline fun dateText(
        format: String,
        excelDate: () -> Double,
    ): String =
        when (format) {
            // Fast path for the default formats: their output has the same length as the pattern
            StyleResolver.DEFAULT_DATE_FORMAT, StyleResolver.DEFAULT_DATE_TIME_FORMAT -> format
            else -> formatter.formatRawCellContents(excelDate(), -1, format)
        }

    private fun plainText(value: Number): String =
        when (value) {
            is Int, is Long, is Short, is Byte, is BigInteger -> value.toString()
            is BigDecimal -> value.stripTrailingZeros().toPlainString()
            is Float -> plainDecimal(value.toString(), value.toDouble())
            else -> plainDecimal(value.toDouble().toString(), value.toDouble())
        }

    private fun plainDecimal(
        text: String,
        value: Double,
    ): String =
        when {
            !value.isFinite() -> text
            value % 1.0 == 0.0 && abs(value) < MAX_EXACT_INTEGRAL_DOUBLE -> value.toLong().toString()
            text.contains('E') -> BigDecimal(text).stripTrailingZeros().toPlainString()
            else -> text
        }
}
