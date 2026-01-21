package io.clroot.excel.parser

import io.clroot.excel.core.ExcelException

/**
 * Result of parsing an Excel file.
 */
sealed class ParseResult<T> {
    /**
     * Successful parse result containing the parsed data.
     */
    data class Success<T>(val data: List<T>) : ParseResult<T>()

    /**
     * Failed parse result containing the list of errors.
     */
    data class Failure<T>(val errors: List<ParseError>) : ParseResult<T>()

    /**
     * Returns the parsed data or throws [ExcelParseException] if parsing failed.
     */
    fun getOrThrow(): List<T> =
        when (this) {
            is Success -> data
            is Failure -> throw ExcelParseException(
                message = "Excel parsing failed with ${errors.size} error(s)",
                errors = errors,
            )
        }

    /**
     * Returns the parsed data or the result of [default] if parsing failed.
     */
    fun getOrElse(default: () -> List<T>): List<T> =
        when (this) {
            is Success -> data
            is Failure -> default()
        }

    /**
     * Returns the parsed data or null if parsing failed.
     */
    fun getOrNull(): List<T>? =
        when (this) {
            is Success -> data
            is Failure -> null
        }
}

/**
 * Represents a single parsing error.
 */
data class ParseError(
    val rowIndex: Int,
    val columnHeader: String?,
    val message: String,
    val cause: Throwable? = null,
)

/**
 * Exception thrown when Excel parsing fails.
 */
class ExcelParseException(
    message: String,
    val errors: List<ParseError>,
    cause: Throwable? = null,
) : ExcelException(message, cause)
