package io.clroot.excel.parser

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.InputStream
import kotlin.reflect.KClass
import kotlin.reflect.KParameter
import kotlin.reflect.full.primaryConstructor

/**
 * Parses an Excel file into a list of data class instances.
 */
inline fun <reified T : Any> parseExcel(
    input: InputStream,
    noinline configure: ParseConfig.Builder<T>.() -> Unit = {},
): ParseResult<T> {
    return parseExcel(T::class, input, configure)
}

/**
 * Parses an Excel file into a list of data class instances.
 */
fun <T : Any> parseExcel(
    klass: KClass<T>,
    input: InputStream,
    configure: ParseConfig.Builder<T>.() -> Unit = {},
): ParseResult<T> {
    val config = ParseConfig.Builder<T>().apply(configure).build()
    return ExcelParserImpl(klass, config).parse(input)
}

/**
 * Internal parser implementation.
 */
internal class ExcelParserImpl<T : Any>(
    private val klass: KClass<T>,
    private val config: ParseConfig<T>,
) {
    private val columns: List<ColumnMeta> = AnnotationExtractor.extractColumns(klass)
    private val headerMatcher: HeaderMatcher = HeaderMatcher(config.headerMatching)
    private val cellConverter: CellConverter =
        CellConverter(
            customConverters = config.converters,
            trimWhitespace = config.trimWhitespace,
        )

    fun parse(input: InputStream): ParseResult<T> {
        val errors = mutableListOf<ParseError>()
        val results = mutableListOf<T>()

        WorkbookFactory.create(input).use { workbook ->
            val sheet = getSheet(workbook)
            val headerMapping = parseHeaders(sheet)

            if (headerMapping == null) {
                return ParseResult.Failure(
                    listOf(
                        ParseError(
                            rowIndex = config.headerRow,
                            columnHeader = null,
                            message = "Failed to parse headers",
                        ),
                    ),
                )
            }

            // Check for missing required columns
            val foundPropertyNames = headerMapping.values.map { it.propertyName }.toSet()
            val missingColumns = columns.filter { !it.isNullable && it.propertyName !in foundPropertyNames }
            if (missingColumns.isNotEmpty()) {
                val missingErrors =
                    missingColumns.map {
                        ParseError(
                            rowIndex = config.headerRow,
                            columnHeader = it.header,
                            message = "필수 컬럼 '${it.header}'을(를) 찾을 수 없습니다.",
                        )
                    }
                return ParseResult.Failure(missingErrors)
            }

            val dataStartRow = config.headerRow + 1
            val lastRowNum = sheet.lastRowNum

            for (rowIndex in dataStartRow..lastRowNum) {
                val row = sheet.getRow(rowIndex)

                if (row == null || isEmptyRow(row)) {
                    if (config.skipEmptyRows) continue
                }

                val parseRowResult = parseRow(row, rowIndex, headerMapping)

                when (parseRowResult) {
                    is RowParseResult.Success -> {
                        // Validate row
                        val validationError = validateRow(parseRowResult.data, rowIndex)
                        if (validationError != null) {
                            errors.add(validationError)
                            if (config.onError == OnError.FAIL_FAST) {
                                return ParseResult.Failure(errors)
                            }
                        } else {
                            results.add(parseRowResult.data)
                        }
                    }

                    is RowParseResult.Failure -> {
                        errors.addAll(parseRowResult.errors)
                        if (config.onError == OnError.FAIL_FAST) {
                            return ParseResult.Failure(errors)
                        }
                    }
                }
            }
        }

        // Validate all
        if (errors.isEmpty() && config.allValidator != null) {
            val allValidationError = validateAll(results)
            if (allValidationError != null) {
                return ParseResult.Failure(listOf(allValidationError))
            }
        }

        return if (errors.isEmpty()) {
            ParseResult.Success(results)
        } else {
            ParseResult.Failure(errors)
        }
    }

    private fun getSheet(workbook: Workbook): Sheet {
        return if (config.sheetName != null) {
            workbook.getSheet(config.sheetName)
                ?: throw IllegalArgumentException("Sheet '${config.sheetName}' not found")
        } else {
            workbook.getSheetAt(config.sheetIndex)
        }
    }

    private fun parseHeaders(sheet: Sheet): Map<Int, ColumnMeta>? {
        val headerRow = sheet.getRow(config.headerRow) ?: return null
        val mapping = mutableMapOf<Int, ColumnMeta>()

        for (cellIndex in 0 until headerRow.lastCellNum) {
            val cell = headerRow.getCell(cellIndex) ?: continue
            val headerValue = getCellStringValue(cell)

            val matchedColumn = headerMatcher.findMatch(headerValue, columns)
            if (matchedColumn != null) {
                mapping[cellIndex] = matchedColumn
            }
        }

        return mapping
    }

    private fun parseRow(
        row: Row?,
        rowIndex: Int,
        headerMapping: Map<Int, ColumnMeta>,
    ): RowParseResult<T> {
        val errors = mutableListOf<ParseError>()
        val values = mutableMapOf<String, Any?>()

        // Initialize with nulls for all columns
        columns.forEach { col ->
            values[col.propertyName] = null
        }

        if (row != null) {
            for ((cellIndex, columnMeta) in headerMapping) {
                val cell = row.getCell(cellIndex)
                val cellValue = getCellValue(cell)

                try {
                    val convertedValue =
                        cellConverter.convert(
                            cellValue,
                            columnMeta.propertyType,
                            columnMeta.isNullable,
                        )
                    values[columnMeta.propertyName] = convertedValue
                } catch (e: Exception) {
                    errors.add(
                        ParseError(
                            rowIndex = rowIndex,
                            columnHeader = columnMeta.header,
                            message =
                                "Failed to convert value '$cellValue' " +
                                    "to ${columnMeta.propertyType.simpleName}: ${e.message}",
                            cause = e,
                        ),
                    )
                }
            }
        }

        if (errors.isNotEmpty()) {
            return RowParseResult.Failure(errors)
        }

        // Create instance
        return try {
            val instance = createInstance(values)
            RowParseResult.Success(instance)
        } catch (e: Exception) {
            RowParseResult.Failure(
                listOf(
                    ParseError(
                        rowIndex = rowIndex,
                        columnHeader = null,
                        message = "Failed to create instance: ${e.message}",
                        cause = e,
                    ),
                ),
            )
        }
    }

    private fun createInstance(values: Map<String, Any?>): T {
        val constructor =
            klass.primaryConstructor
                ?: throw IllegalStateException("No primary constructor found for ${klass.simpleName}")

        val args = mutableMapOf<KParameter, Any?>()
        for (param in constructor.parameters) {
            val value = values[param.name]
            if (value == null && !param.type.isMarkedNullable && !param.isOptional) {
                throw IllegalArgumentException("Missing required value for parameter '${param.name}'")
            }
            if (value != null || param.type.isMarkedNullable) {
                args[param] = value
            }
        }

        return constructor.callBy(args)
    }

    private fun getCellValue(cell: Cell?): Any? {
        if (cell == null) return null

        return when (cell.cellType) {
            CellType.STRING -> cell.stringCellValue
            CellType.NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    cell.localDateTimeCellValue
                } else {
                    cell.numericCellValue
                }
            }

            CellType.BOOLEAN -> cell.booleanCellValue
            CellType.BLANK -> null
            CellType.FORMULA -> {
                when (cell.cachedFormulaResultType) {
                    CellType.STRING -> cell.stringCellValue
                    CellType.NUMERIC -> cell.numericCellValue
                    CellType.BOOLEAN -> cell.booleanCellValue
                    else -> null
                }
            }

            else -> null
        }
    }

    private fun getCellStringValue(cell: Cell): String {
        return when (cell.cellType) {
            CellType.STRING -> cell.stringCellValue
            CellType.NUMERIC -> cell.numericCellValue.toString()
            else -> cell.toString()
        }
    }

    private fun isEmptyRow(row: Row): Boolean {
        for (cellIndex in 0 until row.lastCellNum) {
            val cell = row.getCell(cellIndex)
            if (cell != null && cell.cellType != CellType.BLANK) {
                val value = getCellValue(cell)
                if (value != null && value.toString().isNotBlank()) {
                    return false
                }
            }
        }
        return true
    }

    private fun validateRow(
        data: T,
        rowIndex: Int,
    ): ParseError? {
        val validator = config.rowValidator ?: return null
        return try {
            validator(data)
            null
        } catch (e: Exception) {
            ParseError(
                rowIndex = rowIndex,
                columnHeader = null,
                message = e.message ?: "Row validation failed",
                cause = e,
            )
        }
    }

    private fun validateAll(data: List<T>): ParseError? {
        val validator = config.allValidator ?: return null
        return try {
            validator(data)
            null
        } catch (e: Exception) {
            ParseError(
                rowIndex = -1,
                columnHeader = null,
                message = e.message ?: "Validation failed",
                cause = e,
            )
        }
    }

    private sealed class RowParseResult<T> {
        data class Success<T>(val data: T) : RowParseResult<T>()

        data class Failure<T>(val errors: List<ParseError>) : RowParseResult<T>()
    }
}
