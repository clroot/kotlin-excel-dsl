package io.clroot.excel.render

import io.clroot.excel.core.model.ColumnDefinition
import io.clroot.excel.core.model.ColumnWidth
import io.clroot.excel.core.model.Formula
import io.clroot.excel.core.model.HeaderGroup
import io.clroot.excel.core.model.Sheet
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.streaming.SXSSFSheet
import java.time.LocalDate
import java.time.LocalDateTime

/**
 * Renders a single sheet within an Excel workbook.
 *
 * This class handles:
 * - Header rendering (simple and grouped headers)
 * - Data row rendering with streaming
 * - Column width application
 * - Freeze pane and auto-filter configuration
 *
 * @property sheet the POI sheet to render to
 * @property sheetModel the domain sheet model
 * @property styleResolver resolves styles for cells
 * @property styleCache caches POI styles
 */
internal class SheetRenderer(
    private val sheet: SXSSFSheet,
    private val sheetModel: Sheet,
    private val styleResolver: StyleResolver,
    private val styleCache: StyleCache,
) {
    private val columns = sheetModel.columns
    private val hasHeaderGroups = sheetModel.headerGroups.isNotEmpty()
    private val widthTracker = ColumnWidthTracker(columns)
    private var currentRowIndex = 0

    companion object {
        /** Default column width in POI units (8 characters) */
        private const val DEFAULT_COLUMN_WIDTH = 8 * 256
    }

    /**
     * Renders the complete sheet including headers, data, and configuration.
     */
    fun render() {
        renderHeaders()
        renderDataRows()
        applyColumnWidths()
        applyFreezePane()
        applyAutoFilter()
    }

    private fun renderHeaders() {
        if (hasHeaderGroups) {
            renderGroupedHeaders()
        } else {
            renderSimpleHeader()
        }
    }

    /**
     * Renders a two-row header.
     *
     * Row 1 holds group titles merged across each group's columns, placed at the
     * group's actual position in [columns]. A column that belongs to no group gets
     * its own header text in both rows, merged vertically (e.g. A1:A2). Row 2 always
     * holds every column header, so the auto-filter row and readers of row 2 see all names.
     */
    private fun renderGroupedHeaders() {
        val groupHeaderRowIndex = currentRowIndex++
        val columnHeaderRowIndex = currentRowIndex++
        val groupHeaderRow = sheet.createRow(groupHeaderRowIndex)
        val columnHeaderRow = sheet.createRow(columnHeaderRowIndex)

        val groupedIndices = mutableSetOf<Int>()
        locateHeaderGroups().forEach { (group, startCol) ->
            val endCol = startCol + group.columns.size - 1
            val cell = groupHeaderRow.createCell(startCol)
            cell.setCellValue(group.title)

            // Apply global header style to group headers
            styleResolver.resolveGroupHeaderStyle()?.let {
                cell.cellStyle = styleCache.getOrCreate(it)
            }

            if (endCol > startCol) {
                sheet.addMergedRegion(CellRangeAddress(groupHeaderRowIndex, groupHeaderRowIndex, startCol, endCol))
            }
            groupedIndices.addAll(startCol..endCol)
        }

        columns.forEachIndexed { index, column ->
            createColumnHeaderCell(columnHeaderRow, index, column)
            if (index !in groupedIndices) {
                createColumnHeaderCell(groupHeaderRow, index, column)
                sheet.addMergedRegion(CellRangeAddress(groupHeaderRowIndex, columnHeaderRowIndex, index, index))
            }
        }
    }

    /**
     * Finds the start column index of each header group within [columns].
     *
     * Groups are matched in declaration order by the identity of their first column,
     * falling back to equality for sheets built without the DSL. Groups that cannot
     * be located as a contiguous run of columns are skipped.
     */
    private fun locateHeaderGroups(): List<Pair<HeaderGroup, Int>> {
        var searchFrom = 0
        return sheetModel.headerGroups.mapNotNull { group ->
            if (group.columns.isEmpty()) return@mapNotNull null
            val start =
                findGroupStart(group, searchFrom) { a, b -> a === b }
                    ?: findGroupStart(group, searchFrom) { a, b -> a == b }
                    ?: return@mapNotNull null
            searchFrom = start + group.columns.size
            group to start
        }
    }

    private inline fun findGroupStart(
        group: HeaderGroup,
        searchFrom: Int,
        matches: (ColumnDefinition<*>, ColumnDefinition<*>) -> Boolean,
    ): Int? =
        (searchFrom..columns.size - group.columns.size).firstOrNull { start ->
            group.columns.indices.all { offset -> matches(columns[start + offset], group.columns[offset]) }
        }

    private fun renderSimpleHeader() {
        if (columns.isEmpty()) return

        val headerRow = sheet.createRow(currentRowIndex++)
        columns.forEachIndexed { index, column ->
            createColumnHeaderCell(headerRow, index, column)
        }
    }

    private fun createColumnHeaderCell(
        row: Row,
        index: Int,
        column: ColumnDefinition<*>,
    ) {
        val cell = row.createCell(index)
        cell.setCellValue(column.header)
        styleResolver.resolveHeaderStyle(column)?.let {
            cell.cellStyle = styleCache.getOrCreate(it)
        }
    }

    private fun renderDataRows() {
        val dataSource = sheetModel.dataSource ?: return

        @Suppress("UNCHECKED_CAST")
        val typedColumns = columns as List<ColumnDefinition<Any?>>

        var dataRowIndex = 0
        dataSource.forEach { item ->
            val row = sheet.createRow(currentRowIndex++)
            val isAlternateRow = dataRowIndex % 2 == 0 && sheetModel.alternateRowStyle != null

            typedColumns.forEachIndexed { cellIndex, column ->
                val value = column.valueExtractor(item)
                renderCell(row, cellIndex, column, value, isAlternateRow)
            }

            dataRowIndex++
        }
    }

    private fun renderCell(
        row: Row,
        cellIndex: Int,
        column: ColumnDefinition<Any?>,
        value: Any?,
        isAlternateRow: Boolean,
    ) {
        val cell = row.createCell(cellIndex)

        // Set cell value
        setCellValue(cell, value)

        // Determine date format if applicable
        val dateFormat = getDateFormat(value)

        // Resolve and apply style (pass value for conditional style evaluation)
        val finalStyle = styleResolver.resolveFinalStyle(column, sheetModel, isAlternateRow, dateFormat, value)
        finalStyle?.let { cell.cellStyle = styleCache.getOrCreate(it) }

        // Track width for auto-width columns
        widthTracker.trackWidth(cellIndex, value)
    }

    private fun setCellValue(
        cell: Cell,
        value: Any?,
    ) {
        when (value) {
            null -> cell.setBlank()
            is String -> cell.setCellValue(value)
            is Number -> cell.setCellValue(value.toDouble())
            is Boolean -> cell.setCellValue(value)
            is LocalDate -> cell.setCellValue(value)
            is LocalDateTime -> cell.setCellValue(value)
            is Formula -> cell.cellFormula = value.normalizedExpression
            else -> cell.setCellValue(value.toString())
        }
    }

    private fun getDateFormat(value: Any?): String? {
        return when (value) {
            is LocalDate -> "yyyy-mm-dd"
            is LocalDateTime -> "yyyy-mm-dd hh:mm:ss"
            else -> null
        }
    }

    private fun applyColumnWidths() {
        columns.forEachIndexed { index, column ->
            val width =
                when (val w = column.width) {
                    is ColumnWidth.Fixed -> w.chars * 256
                    is ColumnWidth.Auto -> widthTracker.getPoiWidth(index) ?: DEFAULT_COLUMN_WIDTH
                }
            sheet.setColumnWidth(index, width)
        }
    }

    private fun applyFreezePane() {
        sheetModel.freezePane?.let { freeze ->
            if (freeze.row > 0 || freeze.col > 0) {
                sheet.createFreezePane(freeze.col, freeze.row)
            }
        }
    }

    private fun applyAutoFilter() {
        if (!sheetModel.autoFilter || columns.isEmpty()) return

        val headerRowIndex = if (hasHeaderGroups) 1 else 0
        val lastColIndex = columns.size - 1
        sheet.setAutoFilter(CellRangeAddress(headerRowIndex, headerRowIndex, 0, lastColIndex))
    }
}
