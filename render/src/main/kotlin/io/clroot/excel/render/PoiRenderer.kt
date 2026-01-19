package io.clroot.excel.render

import io.clroot.excel.core.ExcelWriteException
import io.clroot.excel.core.model.*
import io.clroot.excel.core.style.*
import org.apache.poi.ss.usermodel.BorderStyle as PoiBorderStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFColor
import java.io.OutputStream
import java.time.LocalDate
import java.time.LocalDateTime

/**
 * Renders ExcelDocument to .xlsx format using Apache POI (SXSSF for streaming).
 */
class PoiRenderer(
    private val rowAccessWindowSize: Int = 100
) : ExcelRenderer {

    override fun render(document: ExcelDocument, output: OutputStream) {
        try {
            SXSSFWorkbook(rowAccessWindowSize).use { workbook ->
                // Create global styles
                val globalHeaderStyle = document.headerStyle?.let { createPoiStyle(workbook, it) }
                val globalBodyStyle = document.bodyStyle?.let { createPoiStyle(workbook, it) }

                // Create column-specific styles from document-level configuration
                val columnHeaderStyles = mutableMapOf<String, org.apache.poi.ss.usermodel.CellStyle>()
                val columnBodyStyles = mutableMapOf<String, org.apache.poi.ss.usermodel.CellStyle>()
                document.columnStyles.forEach { (columnName, config) ->
                    config.headerStyle?.let { columnHeaderStyles[columnName] = createPoiStyle(workbook, it) }
                    config.bodyStyle?.let { columnBodyStyles[columnName] = createPoiStyle(workbook, it) }
                }

                // Create date style
                val dateStyle = workbook.createCellStyle().apply {
                    dataFormat = workbook.createDataFormat().getFormat("yyyy-mm-dd")
                }
                val dateTimeStyle = workbook.createCellStyle().apply {
                    dataFormat = workbook.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss")
                }

                document.sheets.forEach { sheetModel ->
                    // Create inline column styles (defined in column DSL)
                    val inlineHeaderStyles = mutableMapOf<Int, org.apache.poi.ss.usermodel.CellStyle>()
                    val inlineBodyStyles = mutableMapOf<Int, org.apache.poi.ss.usermodel.CellStyle>()
                    sheetModel.columns.forEachIndexed { index, column ->
                        column.headerStyle?.let { inlineHeaderStyles[index] = createPoiStyle(workbook, it) }
                        column.bodyStyle?.let { inlineBodyStyles[index] = createPoiStyle(workbook, it) }
                    }
                    val sheet = workbook.createSheet(sheetModel.name)
                    var currentRow = 0

                    // Check if we have header groups
                    val hasHeaderGroups = sheetModel.headerGroups.isNotEmpty()

                    // Helper to get header style for a column (priority: inline > column-specific > global)
                    fun getHeaderStyleForColumn(
                        index: Int,
                        columnHeader: String
                    ): org.apache.poi.ss.usermodel.CellStyle? {
                        return inlineHeaderStyles[index]
                            ?: columnHeaderStyles[columnHeader]
                            ?: globalHeaderStyle
                    }

                    // Helper to get body style for a column (priority: inline > column-specific > global)
                    fun getBodyStyleForColumn(
                        index: Int,
                        columnHeader: String
                    ): org.apache.poi.ss.usermodel.CellStyle? {
                        return inlineBodyStyles[index]
                            ?: columnBodyStyles[columnHeader]
                            ?: globalBodyStyle
                    }

                    if (hasHeaderGroups) {
                        // Create group header row (first row)
                        val groupHeaderRow = sheet.createRow(currentRow++)
                        var colIndex = 0

                        sheetModel.headerGroups.forEach { group ->
                            val startCol = colIndex
                            val cell = groupHeaderRow.createCell(startCol)
                            cell.setCellValue(group.title)
                            globalHeaderStyle?.let { cell.cellStyle = it }

                            // Skip columns for this group
                            colIndex += group.columns.size

                            // Merge cells for group header
                            if (group.columns.size > 1) {
                                sheet.addMergedRegion(CellRangeAddress(0, 0, startCol, colIndex - 1))
                            }
                        }

                        // Create column header row (second row)
                        val columnHeaderRow = sheet.createRow(currentRow++)
                        sheetModel.columns.forEachIndexed { index, column ->
                            val cell = columnHeaderRow.createCell(index)
                            cell.setCellValue(column.header)
                            getHeaderStyleForColumn(index, column.header)?.let { cell.cellStyle = it }
                        }
                    } else {
                        // Simple header row
                        if (sheetModel.columns.isNotEmpty()) {
                            val headerRow = sheet.createRow(currentRow++)
                            sheetModel.columns.forEachIndexed { index, column ->
                                val cell = headerRow.createCell(index)
                                cell.setCellValue(column.header)
                                getHeaderStyleForColumn(index, column.header)?.let { cell.cellStyle = it }
                            }
                        }
                    }

                    // Create data rows
                    sheetModel.rows.forEachIndexed { rowIndex, rowModel ->
                        val row = sheet.createRow(currentRow + rowIndex)
                        rowModel.cells.forEachIndexed { cellIndex, cellModel ->
                            val cell = row.createCell(cellIndex)
                            setCellValue(cell, cellModel.value, dateStyle, dateTimeStyle)
                            // Apply body style only for non-date cells (date cells have their own format)
                            if (cellModel.value !is LocalDate && cellModel.value !is LocalDateTime) {
                                val columnHeader = sheetModel.columns.getOrNull(cellIndex)?.header ?: ""
                                getBodyStyleForColumn(cellIndex, columnHeader)?.let { cell.cellStyle = it }
                            }
                        }
                    }

                    // Set column widths
                    sheetModel.columns.forEachIndexed { index, column ->
                        when (val width = column.width) {
                            is ColumnWidth.Fixed -> sheet.setColumnWidth(index, width.chars * 256)
                            is ColumnWidth.Auto -> sheet.trackColumnForAutoSizing(index)
                            is ColumnWidth.Percent -> {} // Handle percentage-based width if needed
                        }
                    }
                }

                workbook.write(output)
            }
        } catch (e: Exception) {
            throw ExcelWriteException(
                message = "Failed to write Excel document: ${e.message}",
                cause = e
            )
        }
    }

    private fun createPoiStyle(workbook: Workbook, style: CellStyle): org.apache.poi.ss.usermodel.CellStyle {
        val poiStyle = workbook.createCellStyle()

        // Background color
        style.backgroundColor?.let { color ->
            poiStyle.fillForegroundColor = createXSSFColor(color).indexed
            poiStyle.setFillForegroundColor(createXSSFColor(color))
            poiStyle.fillPattern = FillPatternType.SOLID_FOREGROUND
        }

        // Font
        val font = workbook.createFont()
        if (style.bold) {
            font.bold = true
        }
        if (style.italic) {
            font.italic = true
        }
        style.fontColor?.let { color ->
            if (font is org.apache.poi.xssf.usermodel.XSSFFont) {
                font.setColor(createXSSFColor(color))
            }
        }
        poiStyle.setFont(font)

        // Alignment
        style.alignment?.let { alignment ->
            poiStyle.alignment = when (alignment) {
                Alignment.LEFT -> HorizontalAlignment.LEFT
                Alignment.CENTER -> HorizontalAlignment.CENTER
                Alignment.RIGHT -> HorizontalAlignment.RIGHT
            }
        }

        // Border
        style.border?.let { border ->
            val poiBorder = when (border) {
                BorderStyle.NONE -> PoiBorderStyle.NONE
                BorderStyle.THIN -> PoiBorderStyle.THIN
                BorderStyle.MEDIUM -> PoiBorderStyle.MEDIUM
                BorderStyle.THICK -> PoiBorderStyle.THICK
            }
            poiStyle.borderTop = poiBorder
            poiStyle.borderBottom = poiBorder
            poiStyle.borderLeft = poiBorder
            poiStyle.borderRight = poiBorder
        }

        // Number format
        style.numberFormat?.let { format ->
            poiStyle.dataFormat = workbook.createDataFormat().getFormat(format)
        }

        return poiStyle
    }

    private fun createXSSFColor(color: Color): XSSFColor {
        return XSSFColor(byteArrayOf(color.red.toByte(), color.green.toByte(), color.blue.toByte()), null)
    }

    private fun setCellValue(
        cell: org.apache.poi.ss.usermodel.Cell,
        value: Any?,
        dateStyle: org.apache.poi.ss.usermodel.CellStyle,
        dateTimeStyle: org.apache.poi.ss.usermodel.CellStyle
    ) {
        when (value) {
            null -> cell.setBlank()
            is String -> cell.setCellValue(value)
            is Number -> cell.setCellValue(value.toDouble())
            is Boolean -> cell.setCellValue(value)
            is LocalDate -> {
                cell.setCellValue(value)
                cell.cellStyle = dateStyle
            }

            is LocalDateTime -> {
                cell.setCellValue(value)
                cell.cellStyle = dateTimeStyle
            }

            else -> cell.setCellValue(value.toString())
        }
    }
}
