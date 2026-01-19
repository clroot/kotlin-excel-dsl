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
)

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
