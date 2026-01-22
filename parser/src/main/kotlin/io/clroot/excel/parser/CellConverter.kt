package io.clroot.excel.parser

import java.math.BigDecimal
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter
import kotlin.reflect.KClass

/**
 * Converts Excel cell values to Kotlin types.
 *
 * Handles conversion from POI cell values (String, Double, Boolean, LocalDateTime)
 * to target Kotlin types including primitives, date/time types, and BigDecimal.
 *
 * Supported target types:
 * - String, Int, Long, Double, Float, Boolean
 * - LocalDate, LocalDateTime
 * - BigDecimal
 * - Custom types via registered converters
 *
 * @property customConverters custom type converters keyed by target type
 * @property trimWhitespace whether to trim whitespace from string values
 * @property treatBlankAsNull whether to treat blank strings as null
 */
class CellConverter(
    private val customConverters: Map<KClass<*>, (Any?) -> Any?> = emptyMap(),
    private val trimWhitespace: Boolean = true,
    private val treatBlankAsNull: Boolean = true,
) {
    companion object {
        // Excel epoch is 1899-12-30 (accounting for Excel's leap year bug)
        private val EXCEL_EPOCH = LocalDate.of(1899, 12, 30)
    }

    /**
     * Converts a cell value to the specified target type.
     *
     * @param T the target type
     * @param value the raw cell value from POI (may be null)
     * @param targetType the KClass of the target type
     * @param isNullable whether the target property accepts null values
     * @return the converted value, or null if nullable and value is null/blank
     * @throws IllegalArgumentException if conversion fails or null is not allowed
     */
    @Suppress("UNCHECKED_CAST")
    fun <T : Any> convert(
        value: Any?,
        targetType: KClass<T>,
        isNullable: Boolean = false,
    ): T? {
        // Check custom converter first
        customConverters[targetType]?.let { converter ->
            return converter(value) as T?
        }

        if (value == null) {
            if (isNullable) {
                return null
            } else {
                throw IllegalArgumentException(
                    "${targetType.simpleName} 타입은 null을 허용하지 않지만, 셀 값이 비어있습니다.",
                )
            }
        }

        val processedValue = if (trimWhitespace && value is String) value.trim() else value

        // Treat blank strings as null if configured
        if (treatBlankAsNull && processedValue is String && processedValue.isBlank()) {
            if (isNullable) {
                return null
            } else {
                throw IllegalArgumentException(
                    "${targetType.simpleName} 타입은 null을 허용하지 않지만, 셀 값이 비어있습니다.",
                )
            }
        }

        return when (targetType) {
            String::class -> convertToString(processedValue)
            Int::class -> convertToInt(processedValue)
            Long::class -> convertToLong(processedValue)
            Double::class -> convertToDouble(processedValue)
            Float::class -> convertToFloat(processedValue)
            Boolean::class -> convertToBoolean(processedValue)
            LocalDate::class -> convertToLocalDate(processedValue)
            LocalDateTime::class -> convertToLocalDateTime(processedValue)
            BigDecimal::class -> convertToBigDecimal(processedValue)
            else -> throw IllegalArgumentException("Unsupported type: ${targetType.simpleName}")
        } as T?
    }

    private fun convertToString(value: Any): String {
        return when (value) {
            is String -> value
            is Double ->
                if (value == value.toLong().toDouble()) {
                    value.toLong().toString()
                } else {
                    value.toString()
                }

            else -> value.toString()
        }
    }

    private fun convertToInt(value: Any): Int {
        return when (value) {
            is Number -> value.toInt()
            is String -> value.toDoubleOrNull()?.toInt() ?: value.toInt()
            else -> throw IllegalArgumentException("Cannot convert ${value::class.simpleName} to Int")
        }
    }

    private fun convertToLong(value: Any): Long {
        return when (value) {
            is Number -> value.toLong()
            is String -> value.toDoubleOrNull()?.toLong() ?: value.toLong()
            else -> throw IllegalArgumentException("Cannot convert ${value::class.simpleName} to Long")
        }
    }

    private fun convertToDouble(value: Any): Double {
        return when (value) {
            is Number -> value.toDouble()
            is String -> value.toDouble()
            else -> throw IllegalArgumentException("Cannot convert ${value::class.simpleName} to Double")
        }
    }

    private fun convertToFloat(value: Any): Float {
        return when (value) {
            is Number -> value.toFloat()
            is String -> value.toFloat()
            else -> throw IllegalArgumentException("Cannot convert ${value::class.simpleName} to Float")
        }
    }

    private fun convertToBoolean(value: Any): Boolean {
        return when (value) {
            is Boolean -> value
            is String -> value.lowercase() == "true"
            is Number -> value.toInt() != 0
            else -> throw IllegalArgumentException("Cannot convert ${value::class.simpleName} to Boolean")
        }
    }

    private fun convertToLocalDate(value: Any): LocalDate {
        return when (value) {
            is LocalDate -> value
            is LocalDateTime -> value.toLocalDate()
            is Number -> EXCEL_EPOCH.plusDays(value.toLong())
            is String -> LocalDate.parse(value, DateTimeFormatter.ISO_LOCAL_DATE)
            else -> throw IllegalArgumentException("Cannot convert ${value::class.simpleName} to LocalDate")
        }
    }

    private fun convertToLocalDateTime(value: Any): LocalDateTime {
        return when (value) {
            is LocalDateTime -> value
            is LocalDate -> value.atStartOfDay()
            is Number -> {
                val days = value.toLong()
                val fraction = value.toDouble() - days
                val secondsInDay = (fraction * 24 * 60 * 60).toLong()
                EXCEL_EPOCH.plusDays(days).atStartOfDay().plusSeconds(secondsInDay)
            }

            is String -> LocalDateTime.parse(value, DateTimeFormatter.ISO_LOCAL_DATE_TIME)
            else -> throw IllegalArgumentException("Cannot convert ${value::class.simpleName} to LocalDateTime")
        }
    }

    private fun convertToBigDecimal(value: Any): BigDecimal {
        return when (value) {
            is BigDecimal -> value
            is Number -> BigDecimal(value.toString())
            is String -> BigDecimal(value)
            else -> throw IllegalArgumentException("Cannot convert ${value::class.simpleName} to BigDecimal")
        }
    }
}
