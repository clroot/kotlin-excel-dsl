package io.clroot.excel.render

/**
 * Calculates column widths considering CJK (Chinese, Japanese, Korean) characters.
 *
 * POI's default auto-sizing does not properly account for CJK characters,
 * which are typically displayed as full-width (taking approximately 2 character widths).
 * This calculator provides accurate width estimation for mixed-content columns.
 */
internal object ColumnWidthCalculator {
    /**
     * Default character width multiplier for CJK characters.
     * CJK characters are typically 1.5-2x the width of ASCII characters.
     */
    private const val CJK_WIDTH_MULTIPLIER = 2.0

    /**
     * Minimum column width in characters.
     */
    private const val MIN_WIDTH_CHARS = 8

    /**
     * Maximum column width in characters.
     */
    private const val MAX_WIDTH_CHARS = 100

    /**
     * Padding to add to calculated width (in characters).
     */
    private const val PADDING_CHARS = 2

    /**
     * Calculates the optimal width for a column based on its content.
     *
     * @param values the cell values in the column (including header)
     * @return the width in POI units (1/256th of a character)
     */
    fun calculateWidth(values: List<Any?>): Int {
        if (values.isEmpty()) {
            return MIN_WIDTH_CHARS * 256
        }

        val maxWidth =
            values.maxOf { value ->
                calculateTextWidth(value?.toString() ?: "")
            }

        val finalWidth =
            (maxWidth + PADDING_CHARS)
                .coerceIn(MIN_WIDTH_CHARS.toDouble(), MAX_WIDTH_CHARS.toDouble())

        return (finalWidth * 256).toInt()
    }

    /**
     * Calculates the display width of a text string.
     *
     * @param text the text to measure
     * @return the width in character units (considering CJK characters)
     */
    fun calculateTextWidth(text: String): Double {
        if (text.isEmpty()) return 0.0

        var width = 0.0
        for (char in text) {
            width +=
                if (isCjkCharacter(char)) {
                    CJK_WIDTH_MULTIPLIER
                } else {
                    1.0
                }
        }
        return width
    }

    /**
     * Checks if a character is a CJK (Chinese, Japanese, Korean) character.
     *
     * This includes:
     * - CJK Unified Ideographs (U+4E00 - U+9FFF)
     * - CJK Unified Ideographs Extension A (U+3400 - U+4DBF)
     * - Hangul Syllables (U+AC00 - U+D7AF)
     * - Hangul Jamo (U+1100 - U+11FF)
     * - Hangul Compatibility Jamo (U+3130 - U+318F) - includes ㄱ, ㄴ, etc.
     * - Hiragana (U+3040 - U+309F)
     * - Katakana (U+30A0 - U+30FF)
     * - Full-width ASCII variants (U+FF00 - U+FFEF)
     *
     * @param char the character to check
     * @return true if the character is a CJK character
     */
    fun isCjkCharacter(char: Char): Boolean {
        val code = char.code
        return when {
            // CJK Unified Ideographs
            code in 0x4E00..0x9FFF -> true
            // CJK Unified Ideographs Extension A
            code in 0x3400..0x4DBF -> true
            // Hangul Syllables (Korean) - 가 to 힣
            code in 0xAC00..0xD7AF -> true
            // Hangul Jamo
            code in 0x1100..0x11FF -> true
            // Hangul Compatibility Jamo (ㄱ, ㄴ, ㄷ, etc.)
            code in 0x3130..0x318F -> true
            // Hiragana (Japanese)
            code in 0x3040..0x309F -> true
            // Katakana (Japanese)
            code in 0x30A0..0x30FF -> true
            // Full-width forms
            code in 0xFF00..0xFFEF -> true
            // CJK Symbols and Punctuation
            code in 0x3000..0x303F -> true
            else -> false
        }
    }
}
