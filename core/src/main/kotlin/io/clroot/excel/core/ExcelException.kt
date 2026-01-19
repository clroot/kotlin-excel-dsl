package io.clroot.excel.core

/**
 * Base exception for all Excel-related errors.
 */
open class ExcelException(
    message: String,
    cause: Throwable? = null,
) : RuntimeException(message, cause)

/**
 * Exception thrown when there's a problem with the data being processed.
 */
class ExcelDataException(
    message: String,
    sheetName: String? = null,
    rowIndex: Int? = null,
    columnIndex: Int? = null,
    columnHeader: String? = null,
    actualValue: Any? = null,
    cause: Throwable? = null,
) : ExcelException(buildDataErrorMessage(message, sheetName, rowIndex, columnIndex, columnHeader, actualValue), cause)

/**
 * Exception thrown when there's an error writing the Excel file.
 */
class ExcelWriteException(
    message: String,
    sheetName: String? = null,
    rowIndex: Int? = null,
    columnIndex: Int? = null,
    cause: Throwable? = null,
) : ExcelException(buildWriteErrorMessage(message, sheetName, rowIndex, columnIndex), cause)

/**
 * Exception thrown when there's a configuration problem (e.g., annotation parsing errors).
 */
class ExcelConfigurationException(
    message: String,
    className: String? = null,
    propertyName: String? = null,
    hint: String? = null,
    cause: Throwable? = null,
) : ExcelException(buildConfigErrorMessage(message, className, propertyName, hint), cause)

/**
 * Exception thrown when a required column is not found.
 */
class ColumnNotFoundException(
    columnName: String,
    availableColumns: List<String> = emptyList(),
    sheetName: String? = null,
) : ExcelException(buildColumnNotFoundMessage(columnName, availableColumns, sheetName))

/**
 * Exception thrown when a style cannot be applied.
 */
class StyleException(
    message: String,
    styleName: String? = null,
    targetColumn: String? = null,
    cause: Throwable? = null,
) : ExcelException(buildStyleErrorMessage(message, styleName, targetColumn), cause)

// Helper functions to build detailed error messages
private fun buildDataErrorMessage(
    message: String,
    sheetName: String?,
    rowIndex: Int?,
    columnIndex: Int?,
    columnHeader: String?,
    actualValue: Any?,
): String =
    buildString {
        append(message)
        val context = mutableListOf<String>()
        sheetName?.let { context.add("sheet='$it'") }
        rowIndex?.let { context.add("row=${it + 1}") }
        columnIndex?.let { context.add("column=${it + 1}") }
        columnHeader?.let { context.add("header='$it'") }
        actualValue?.let { context.add("value='$it' (${it::class.simpleName})") }
        if (context.isNotEmpty()) {
            append(" [")
            append(context.joinToString(", "))
            append("]")
        }
    }

private fun buildWriteErrorMessage(
    message: String,
    sheetName: String?,
    rowIndex: Int?,
    columnIndex: Int?,
): String =
    buildString {
        append(message)
        val context = mutableListOf<String>()
        sheetName?.let { context.add("sheet='$it'") }
        rowIndex?.let { context.add("row=${it + 1}") }
        columnIndex?.let { context.add("column=${it + 1}") }
        if (context.isNotEmpty()) {
            append(" [")
            append(context.joinToString(", "))
            append("]")
        }
    }

private fun buildConfigErrorMessage(
    message: String,
    className: String?,
    propertyName: String?,
    hint: String?,
): String =
    buildString {
        append(message)
        val context = mutableListOf<String>()
        className?.let { context.add("class=$it") }
        propertyName?.let { context.add("property=$it") }
        if (context.isNotEmpty()) {
            append(" [")
            append(context.joinToString(", "))
            append("]")
        }
        hint?.let {
            append("\nHint: ")
            append(it)
        }
    }

private fun buildColumnNotFoundMessage(
    columnName: String,
    availableColumns: List<String>,
    sheetName: String?,
): String =
    buildString {
        append("Column '$columnName' not found")
        sheetName?.let { append(" (sheet='$it')") }
        if (availableColumns.isNotEmpty()) {
            append("\nAvailable columns: ")
            append(availableColumns.joinToString(", ") { "'$it'" })
        }
    }

private fun buildStyleErrorMessage(
    message: String,
    styleName: String?,
    targetColumn: String?,
): String =
    buildString {
        append(message)
        val context = mutableListOf<String>()
        styleName?.let { context.add("style=$it") }
        targetColumn?.let { context.add("targetColumn='$it'") }
        if (context.isNotEmpty()) {
            append(" [")
            append(context.joinToString(", "))
            append("]")
        }
    }
