package io.clroot.excel.annotation.style

import io.clroot.excel.annotation.BodyStyle
import io.clroot.excel.annotation.HeaderStyle
import io.clroot.excel.core.style.CellStyle
import io.clroot.excel.core.style.Color

/**
 * Utility for converting style annotations to CellStyle.
 */
object StyleConverter {
    /**
     * Parses a HEX color string to Color.
     *
     * @param hex color in "#RRGGBB" or "RRGGBB" format
     * @return parsed Color, or null if empty string
     * @throws IllegalArgumentException if format is invalid
     */
    fun parseHexColor(hex: String): Color? {
        if (hex.isBlank()) return null

        val normalized = hex.removePrefix("#")
        require(normalized.length == 6) { "Invalid HEX color format: $hex" }

        val r = normalized.substring(0, 2).toInt(16)
        val g = normalized.substring(2, 4).toInt(16)
        val b = normalized.substring(4, 6).toInt(16)

        return Color(r, g, b)
    }

    /**
     * Resolves color from enum and hex, preferring hex if not blank.
     */
    fun resolveColor(
        enumColor: StyleColor,
        hexColor: String,
    ): Color? {
        if (hexColor.isNotBlank()) {
            return parseHexColor(hexColor)
        }
        return enumColor.toColor()
    }

    /**
     * Converts @HeaderStyle annotation to CellStyle.
     *
     * @return CellStyle, or null if all values are defaults
     */
    fun toCellStyle(annotation: HeaderStyle): CellStyle? =
        buildCellStyle(
            bg = resolveColor(annotation.backgroundColor, annotation.backgroundColorHex),
            fg = resolveColor(annotation.fontColor, annotation.fontColorHex),
            bold = annotation.bold,
            italic = annotation.italic,
            align = annotation.alignment.toAlignment(),
            border = annotation.border.toBorderStyle(),
        )

    /**
     * Converts @BodyStyle annotation to CellStyle.
     *
     * @return CellStyle, or null if all values are defaults
     */
    fun toCellStyle(annotation: BodyStyle): CellStyle? =
        buildCellStyle(
            bg = resolveColor(annotation.backgroundColor, annotation.backgroundColorHex),
            fg = resolveColor(annotation.fontColor, annotation.fontColorHex),
            bold = annotation.bold,
            italic = annotation.italic,
            align = annotation.alignment.toAlignment(),
            border = annotation.border.toBorderStyle(),
        )

    private fun buildCellStyle(
        bg: Color?,
        fg: Color?,
        bold: Boolean,
        italic: Boolean,
        align: io.clroot.excel.core.style.Alignment?,
        border: io.clroot.excel.core.style.BorderStyle?,
    ): CellStyle? {
        if (bg == null && fg == null && !bold && !italic && align == null && border == null) {
            return null
        }
        return CellStyle(
            backgroundColor = bg,
            fontColor = fg,
            bold = bold,
            italic = italic,
            alignment = align,
            border = border,
        )
    }
}
