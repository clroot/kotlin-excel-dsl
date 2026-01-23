package io.clroot.excel.annotation.style

import io.clroot.excel.core.style.Alignment
import io.clroot.excel.core.style.BorderStyle
import io.clroot.excel.core.style.Color

/**
 * Predefined color constants.
 * For custom colors, use `backgroundColorHex = "#RRGGBB"`.
 */
enum class StyleColor(internal val rgb: Triple<Int, Int, Int>?) {
    NONE(null),
    WHITE(Triple(255, 255, 255)),
    BLACK(Triple(0, 0, 0)),
    GRAY(Triple(128, 128, 128)),
    LIGHT_GRAY(Triple(211, 211, 211)),
    RED(Triple(255, 0, 0)),
    GREEN(Triple(0, 128, 0)),
    BLUE(Triple(0, 0, 255)),
    YELLOW(Triple(255, 255, 0)),
    ORANGE(Triple(255, 165, 0)),
    ;

    internal fun toColor(): Color? = rgb?.let { Color(it.first, it.second, it.third) }
}

/**
 * Text alignment options.
 */
enum class StyleAlignment {
    NONE,
    LEFT,
    CENTER,
    RIGHT,
    ;

    internal fun toAlignment(): Alignment? =
        when (this) {
            NONE -> null
            LEFT -> Alignment.LEFT
            CENTER -> Alignment.CENTER
            RIGHT -> Alignment.RIGHT
        }
}

/**
 * Border style options.
 */
enum class StyleBorder {
    NONE,
    THIN,
    MEDIUM,
    THICK,
    ;

    internal fun toBorderStyle(): BorderStyle? =
        when (this) {
            NONE -> null
            THIN -> BorderStyle.THIN
            MEDIUM -> BorderStyle.MEDIUM
            THICK -> BorderStyle.THICK
        }
}
