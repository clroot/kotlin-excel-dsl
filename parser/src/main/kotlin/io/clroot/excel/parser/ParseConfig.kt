package io.clroot.excel.parser

import kotlin.reflect.KClass

/**
 * Header matching strategy.
 */
enum class HeaderMatching {
    /** Exact match required */
    EXACT,

    /** Flexible match: trim whitespace, normalize spaces, ignore case */
    FLEXIBLE,
}

/**
 * Error handling strategy.
 */
enum class OnError {
    /** Collect all errors and return them in Failure */
    COLLECT,

    /** Stop at first error */
    FAIL_FAST,
}

/**
 * Configuration for Excel parsing.
 */
data class ParseConfig<T : Any>(
    val headerRow: Int,
    val sheetIndex: Int,
    val sheetName: String?,
    val headerMatching: HeaderMatching,
    val onError: OnError,
    val skipEmptyRows: Boolean,
    val trimWhitespace: Boolean,
    val treatBlankAsNull: Boolean,
    val converters: Map<KClass<*>, (Any?) -> Any?>,
    val rowValidator: ((T) -> Unit)?,
    val allValidator: ((List<T>) -> Unit)?,
) {
    /**
     * Builder for constructing [ParseConfig] instances.
     *
     * Example:
     * ```kotlin
     * parseConfig<User> {
     *     headerRow = 1
     *     sheetName = "Users"
     *     headerMatching = HeaderMatching.FLEXIBLE
     *     validateRow { require(it.age > 0) { "Age must be positive" } }
     * }
     * ```
     *
     * @param T the type of data class to parse into
     */
    class Builder<T : Any> {
        var headerRow: Int = 0
        var sheetIndex: Int = 0
        var sheetName: String? = null
        var headerMatching: HeaderMatching = HeaderMatching.FLEXIBLE
        var onError: OnError = OnError.COLLECT
        var skipEmptyRows: Boolean = true
        var trimWhitespace: Boolean = true

        /** Treat blank strings (whitespace only) as null */
        var treatBlankAsNull: Boolean = true

        @PublishedApi
        internal val converters: MutableMap<KClass<*>, (Any?) -> Any?> = mutableMapOf()
        private var rowValidator: ((T) -> Unit)? = null
        private var allValidator: ((List<T>) -> Unit)? = null

        /**
         * Register a custom type converter.
         */
        inline fun <reified R : Any> converter(noinline convert: (Any?) -> R) {
            converters[R::class] = convert
        }

        /**
         * Register a row-level validator.
         */
        fun validateRow(validator: (T) -> Unit) {
            rowValidator = validator
        }

        /**
         * Register a validator for all parsed data.
         */
        fun validateAll(validator: (List<T>) -> Unit) {
            allValidator = validator
        }

        fun build(): ParseConfig<T> =
            ParseConfig(
                headerRow = headerRow,
                sheetIndex = sheetIndex,
                sheetName = sheetName,
                headerMatching = headerMatching,
                onError = onError,
                skipEmptyRows = skipEmptyRows,
                trimWhitespace = trimWhitespace,
                treatBlankAsNull = treatBlankAsNull,
                converters = converters.toMap(),
                rowValidator = rowValidator,
                allValidator = allValidator,
            )
    }
}

/**
 * Creates a new ParseConfig using a DSL builder pattern.
 */
inline fun <reified T : Any> parseConfig(block: ParseConfig.Builder<T>.() -> Unit): ParseConfig<T> =
    ParseConfig.Builder<T>().apply(block).build()
