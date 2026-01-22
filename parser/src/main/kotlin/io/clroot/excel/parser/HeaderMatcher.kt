package io.clroot.excel.parser

import kotlin.reflect.KClass

/**
 * Matches Excel header cells to column definitions.
 */
class HeaderMatcher(private val strategy: HeaderMatching) {
    companion object {
        private val WHITESPACE_REGEX = Regex("\\s+")
    }

    /**
     * Checks if the cell header matches the expected header or any of its aliases.
     */
    fun matches(
        cellHeader: String,
        expectedHeader: String,
        aliases: Array<String>,
    ): Boolean {
        val candidates = listOf(expectedHeader) + aliases
        return candidates.any { candidate ->
            when (strategy) {
                HeaderMatching.EXACT -> cellHeader == candidate
                HeaderMatching.FLEXIBLE -> normalize(cellHeader) == normalize(candidate)
            }
        }
    }

    /**
     * Finds the matching column header from candidates.
     * Returns the matched candidate or null if no match found.
     */
    fun findMatch(
        cellHeader: String,
        columns: List<ColumnMeta>,
    ): ColumnMeta? {
        return columns.find { column ->
            matches(cellHeader, column.header, column.aliases)
        }
    }

    private fun normalize(value: String): String {
        return value
            .trim()
            .replace(WHITESPACE_REGEX, " ")
            .lowercase()
    }
}

/**
 * Metadata extracted from @Column annotation.
 *
 * Note: [equals] and [hashCode] are intentionally implemented based on [propertyName] only.
 * This ensures that a `Set<ColumnMeta>` will correctly identify columns as unique based on their
 * corresponding property, regardless of other metadata like [header] or [aliases].
 */
data class ColumnMeta(
    val header: String,
    val aliases: Array<String>,
    val propertyName: String,
    val propertyType: KClass<*>,
    val isNullable: Boolean,
) {
    override fun equals(other: Any?): Boolean {
        if (this === other) return true
        if (other !is ColumnMeta) return false
        return propertyName == other.propertyName
    }

    override fun hashCode(): Int = propertyName.hashCode()
}
