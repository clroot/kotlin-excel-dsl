package io.clroot.excel.render

import io.clroot.excel.core.ExcelWriteException
import io.clroot.excel.core.model.*
import io.clroot.excel.core.style.*
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFColor
import java.io.OutputStream
import java.time.LocalDate
import java.time.LocalDateTime
import org.apache.poi.ss.usermodel.BorderStyle as PoiBorderStyle

/**
 * Renders ExcelDocument to .xlsx format using Apache POI (SXSSF for streaming).
 *
 * This renderer uses Apache POI's SXSSF (Streaming Usermodel API) for memory-efficient
 * handling of large datasets. It includes style caching to avoid exceeding POI's
 * style limit (approximately 64,000 styles per workbook).
 *
 * @property rowAccessWindowSize the number of rows to keep in memory (default: 100)
 */
class PoiRenderer(
    private val rowAccessWindowSize: Int = 100,
) : ExcelRenderer {
    override fun render(
        document: ExcelDocument,
        output: OutputStream,
    ) {
        try {
            SXSSFWorkbook(rowAccessWindowSize).use { workbook ->
                // Style cache to prevent duplicate style creation
                val styleCache = StyleCache(workbook)

                // Create global styles
                val globalHeaderStyle = document.headerStyle?.let { styleCache.getOrCreate(it) }
                val globalBodyStyle = document.bodyStyle?.let { styleCache.getOrCreate(it) }

                // Create column-specific styles from document-level configuration
                val columnHeaderStyles = mutableMapOf<String, org.apache.poi.ss.usermodel.CellStyle>()
                val columnBodyStyles = mutableMapOf<String, org.apache.poi.ss.usermodel.CellStyle>()
                document.columnStyles.forEach { (columnName, config) ->
                    config.headerStyle?.let { columnHeaderStyles[columnName] = styleCache.getOrCreate(it) }
                    config.bodyStyle?.let { columnBodyStyles[columnName] = styleCache.getOrCreate(it) }
                }

                // Create date style (cached via StyleCache)
                val dateStyle = styleCache.getDateStyle()
                val dateTimeStyle = styleCache.getDateTimeStyle()

                document.sheets.forEach { sheetModel ->
                    // Create inline column styles (defined in column DSL)
                    val inlineHeaderStyles = mutableMapOf<Int, org.apache.poi.ss.usermodel.CellStyle>()
                    val inlineBodyStyles = mutableMapOf<Int, org.apache.poi.ss.usermodel.CellStyle>()
                    sheetModel.columns.forEachIndexed { index, column ->
                        column.headerStyle?.let { inlineHeaderStyles[index] = styleCache.getOrCreate(it) }
                        column.bodyStyle?.let { inlineBodyStyles[index] = styleCache.getOrCreate(it) }
                    }
                    val sheet = workbook.createSheet(sheetModel.name)
                    var currentRow = 0

                    // Check if we have header groups
                    val hasHeaderGroups = sheetModel.headerGroups.isNotEmpty()

                    // Helper to get header style for a column (priority: inline > column-specific > global)
                    fun getHeaderStyleForColumn(
                        index: Int,
                        columnHeader: String,
                    ): org.apache.poi.ss.usermodel.CellStyle? {
                        return inlineHeaderStyles[index]
                            ?: columnHeaderStyles[columnHeader]
                            ?: globalHeaderStyle
                    }

                    // Helper to get body style for a column (priority: inline > column-specific > global)
                    fun getBodyStyleForColumn(
                        index: Int,
                        columnHeader: String,
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
                    // Collect auto-width column indices first to avoid multiple row iterations
                    val autoWidthColumns =
                        sheetModel.columns.mapIndexedNotNull { index, column ->
                            if (column.width is ColumnWidth.Auto) index else null
                        }

                    // Initialize value collectors with headers
                    val autoWidthValues =
                        autoWidthColumns.associateWith { index ->
                            mutableListOf<Any?>(sheetModel.columns[index].header)
                        }

                    // Single pass through rows to collect all auto-width column values
                    if (autoWidthColumns.isNotEmpty()) {
                        sheetModel.rows.forEach { row ->
                            autoWidthColumns.forEach { index ->
                                row.cells.getOrNull(index)?.value?.let { autoWidthValues[index]?.add(it) }
                            }
                        }
                    }

                    // Apply column widths
                    sheetModel.columns.forEachIndexed { index, column ->
                        when (val width = column.width) {
                            is ColumnWidth.Fixed -> sheet.setColumnWidth(index, width.chars * 256)
                            is ColumnWidth.Auto -> {
                                val values = autoWidthValues[index] ?: emptyList()
                                val calculatedWidth = ColumnWidthCalculator.calculateWidth(values)
                                sheet.setColumnWidth(index, calculatedWidth)
                            }

                            is ColumnWidth.Percent -> {} // Handle percentage-based width if needed
                        }
                    }
                }

                workbook.write(output)
            }
        } catch (e: Exception) {
            throw ExcelWriteException(
                message = "Failed to write Excel document: ${e.message}",
                cause = e,
            )
        }
    }

    private fun setCellValue(
        cell: org.apache.poi.ss.usermodel.Cell,
        value: Any?,
        dateStyle: org.apache.poi.ss.usermodel.CellStyle,
        dateTimeStyle: org.apache.poi.ss.usermodel.CellStyle,
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

/**
 * Caches POI CellStyle objects to prevent duplicate creation.
 *
 * POI workbooks have a limit of approximately 64,000 unique cell styles.
 * This cache ensures that identical [CellStyle] configurations are reused,
 * significantly reducing the number of POI styles created.
 *
 * @property workbook the POI workbook to create styles for
 */
internal class StyleCache(private val workbook: Workbook) {
    private val cache = mutableMapOf<CellStyle, org.apache.poi.ss.usermodel.CellStyle>()
    private val fontCache = mutableMapOf<FontKey, org.apache.poi.ss.usermodel.Font>()
    private var dateStyle: org.apache.poi.ss.usermodel.CellStyle? = null
    private var dateTimeStyle: org.apache.poi.ss.usermodel.CellStyle? = null

    /**
     * Gets or creates a POI CellStyle for the given domain CellStyle.
     *
     * If an identical style already exists in the cache, it is returned.
     * Otherwise, a new POI style is created, cached, and returned.
     *
     * @param style the domain CellStyle to convert
     * @return the corresponding POI CellStyle
     */
    fun getOrCreate(style: CellStyle): org.apache.poi.ss.usermodel.CellStyle {
        return cache.getOrPut(style) { createPoiStyle(style) }
    }

    /**
     * Gets or creates a date format style (yyyy-mm-dd).
     */
    fun getDateStyle(): org.apache.poi.ss.usermodel.CellStyle {
        return dateStyle ?: workbook.createCellStyle().apply {
            dataFormat = workbook.createDataFormat().getFormat("yyyy-mm-dd")
        }.also { dateStyle = it }
    }

    /**
     * Gets or creates a datetime format style (yyyy-mm-dd hh:mm:ss).
     */
    fun getDateTimeStyle(): org.apache.poi.ss.usermodel.CellStyle {
        return dateTimeStyle ?: workbook.createCellStyle().apply {
            dataFormat = workbook.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss")
        }.also { dateTimeStyle = it }
    }

    private fun createPoiStyle(style: CellStyle): org.apache.poi.ss.usermodel.CellStyle {
        val poiStyle = workbook.createCellStyle()

        // Background color
        style.backgroundColor?.let { color ->
            poiStyle.fillForegroundColor = createXSSFColor(color).indexed
            poiStyle.setFillForegroundColor(createXSSFColor(color))
            poiStyle.fillPattern = FillPatternType.SOLID_FOREGROUND
        }

        // Font (cached separately)
        val fontKey =
            FontKey(
                bold = style.bold,
                italic = style.italic,
                fontColor = style.fontColor,
            )
        val font = fontCache.getOrPut(fontKey) { createFont(fontKey) }
        poiStyle.setFont(font)

        // Alignment
        style.alignment?.let { alignment ->
            poiStyle.alignment =
                when (alignment) {
                    Alignment.LEFT -> HorizontalAlignment.LEFT
                    Alignment.CENTER -> HorizontalAlignment.CENTER
                    Alignment.RIGHT -> HorizontalAlignment.RIGHT
                }
        }

        // Border
        style.border?.let { border ->
            val poiBorder =
                when (border) {
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

    private fun createFont(fontKey: FontKey): org.apache.poi.ss.usermodel.Font {
        val font = workbook.createFont()
        if (fontKey.bold) {
            font.bold = true
        }
        if (fontKey.italic) {
            font.italic = true
        }
        fontKey.fontColor?.let { color ->
            if (font is org.apache.poi.xssf.usermodel.XSSFFont) {
                font.setColor(createXSSFColor(color))
            }
        }
        return font
    }

    private fun createXSSFColor(color: Color): XSSFColor {
        return XSSFColor(byteArrayOf(color.red.toByte(), color.green.toByte(), color.blue.toByte()), null)
    }

    /**
     * Key for font caching.
     */
    private data class FontKey(
        val bold: Boolean,
        val italic: Boolean,
        val fontColor: Color?,
    )
}
