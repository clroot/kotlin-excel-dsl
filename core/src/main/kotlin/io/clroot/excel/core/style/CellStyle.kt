@file:Suppress("unused")

package io.clroot.excel.core.style

/**
 * Represents cell styling configuration.
 *
 * Every property is optional. `null` means "not specified": the property is inherited from a
 * lower-priority style when styles are [merged][merge], and the Excel default is used when
 * no style specifies it. For [bold] and [italic], `false` explicitly turns the font style off,
 * which lets a higher-priority style override a lower-priority `true`.
 *
 * @property backgroundColor solid background color
 * @property fontColor font color
 * @property bold `true` for bold, `false` for regular weight, `null` to inherit
 * @property italic `true` for italic, `false` for upright, `null` to inherit
 * @property alignment horizontal alignment
 * @property border border style for all four sides
 * @property numberFormat Excel number format code, such as `"#,##0"` or `"yyyy-mm-dd"`
 */
data class CellStyle(
    val backgroundColor: Color? = null,
    val fontColor: Color? = null,
    val bold: Boolean? = null,
    val italic: Boolean? = null,
    val alignment: Alignment? = null,
    val border: BorderStyle? = null,
    val numberFormat: String? = null,
) {
    /**
     * Binary-compatibility constructor for code compiled against versions where
     * [bold] and [italic] were non-null. `false` maps to "not specified" (`null`),
     * matching the previous merge behavior in which `false` never overrode `true`.
     */
    @Deprecated("Kept for binary compatibility only", level = DeprecationLevel.HIDDEN)
    constructor(
        backgroundColor: Color? = null,
        fontColor: Color? = null,
        bold: Boolean = false,
        italic: Boolean = false,
        alignment: Alignment? = null,
        border: BorderStyle? = null,
        numberFormat: String? = null,
    ) : this(
        backgroundColor = backgroundColor,
        fontColor = fontColor,
        bold = bold.takeIf { it },
        italic = italic.takeIf { it },
        alignment = alignment,
        border = border,
        numberFormat = numberFormat,
    )

    /**
     * Merges this style with another style.
     *
     * Every property that is set (non-null) in [other] overrides this style's value;
     * properties that [other] leaves `null` keep this style's value. This also applies to
     * [bold] and [italic], so `other` can turn them off with `false`.
     *
     * @param other the higher-priority style to merge on top of this style
     * @return a new CellStyle with merged properties
     */
    fun merge(other: CellStyle): CellStyle =
        CellStyle(
            backgroundColor = other.backgroundColor ?: this.backgroundColor,
            fontColor = other.fontColor ?: this.fontColor,
            bold = other.bold ?: this.bold,
            italic = other.italic ?: this.italic,
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
