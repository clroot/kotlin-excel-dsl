@file:Suppress("unused")

package io.clroot.excel.core.style

/**
 * Represents cell styling configuration.
 */
data class CellStyle(
    val backgroundColor: Color? = null,
    val fontColor: Color? = null,
    val bold: Boolean = false,
    val italic: Boolean = false,
    val alignment: Alignment? = null,
    val border: BorderStyle? = null,
    val numberFormat: String? = null,
) {
    /**
     * Merges this style with another style.
     * Properties from [other] take precedence when both are set,
     * except for boolean properties which use OR logic.
     *
     * @param other the style to merge with (its properties override this style's)
     * @return a new CellStyle with merged properties
     */
    fun merge(other: CellStyle): CellStyle =
        CellStyle(
            backgroundColor = other.backgroundColor ?: this.backgroundColor,
            fontColor = other.fontColor ?: this.fontColor,
            bold = this.bold || other.bold,
            italic = this.italic || other.italic,
            alignment = other.alignment ?: this.alignment,
            border = other.border ?: this.border,
            numberFormat = other.numberFormat ?: this.numberFormat,
        )
}

/**
 * Color representation for styling.
 */
data class Color(
    val red: Int,
    val green: Int,
    val blue: Int,
) {
    companion object {
        val WHITE = Color(255, 255, 255)
        val BLACK = Color(0, 0, 0)
        val GRAY = Color(128, 128, 128)
        val LIGHT_GRAY = Color(211, 211, 211)
        val RED = Color(255, 0, 0)
        val GREEN = Color(0, 128, 0)
        val BLUE = Color(0, 0, 255)
    }
}

/**
 * Text alignment options.
 */
enum class Alignment {
    LEFT,
    CENTER,
    RIGHT,
}

/**
 * Border style options.
 */
enum class BorderStyle {
    NONE,
    THIN,
    MEDIUM,
    THICK,
}
