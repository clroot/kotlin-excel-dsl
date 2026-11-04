package io.clroot.excel.annotation

import io.clroot.excel.annotation.style.StyleAlignment
import io.clroot.excel.annotation.style.StyleBorder
import io.clroot.excel.annotation.style.StyleColor
import kotlin.reflect.KClass

/**
 * Marks a class as an Excel-exportable data class.
 */
@Target(AnnotationTarget.CLASS)
@Retention(AnnotationRetention.RUNTIME)
annotation class Excel

/**
 * Configures a property as an Excel column.
 *
 * @param header The column header text
 * @param width Column width in characters (0 = auto)
 * @param format Excel number format code applied to body cells, such as `"#,##0"` or `"yyyy/mm/dd"`
 *   (empty = none). It takes precedence over style number formats; a conditional style's
 *   `numberFormat` overrides it. Date cells without a format use the default date format.
 * @param order Column order (lower values appear first)
 * @param aliases Alternative header names for matching during Excel parsing
 */
@Target(AnnotationTarget.PROPERTY)
@Retention(AnnotationRetention.RUNTIME)
annotation class Column(
    val header: String,
    val width: Int = 0,
    val format: String = "",
    val order: Int = Int.MAX_VALUE,
    val aliases: Array<String> = [],
)

/**
 * Defines the style for header cells.
 *
 * When applied at class level, the style applies to all header cells.
 * When applied at property level, it is merged over the class-level style for that column:
 * properties set on the property win, and the others are inherited.
 *
 * If both enum color and hex color are specified, hex takes precedence.
 *
 * @param bold `true` makes the text bold. `false` (default) means "not specified": the value is
 *   inherited from lower-priority styles, so it cannot turn off bold from a theme or class-level
 *   annotation. Use a [ConditionalStyler] returning `CellStyle(bold = false)` for that.
 * @param italic `true` makes the text italic. `false` (default) means "not specified", as for [bold].
 * @param backgroundColor predefined background color
 * @param backgroundColorHex custom background color in "#RRGGBB" format (takes precedence)
 * @param fontColor predefined font color
 * @param fontColorHex custom font color in "#RRGGBB" format (takes precedence)
 * @param alignment text alignment within the cell
 * @param border border style for the cell
 */
@Target(AnnotationTarget.CLASS, AnnotationTarget.PROPERTY)
@Retention(AnnotationRetention.RUNTIME)
annotation class HeaderStyle(
    val bold: Boolean = false,
    val italic: Boolean = false,
    val backgroundColor: StyleColor = StyleColor.NONE,
    val backgroundColorHex: String = "",
    val fontColor: StyleColor = StyleColor.NONE,
    val fontColorHex: String = "",
    val alignment: StyleAlignment = StyleAlignment.NONE,
    val border: StyleBorder = StyleBorder.NONE,
)

/**
 * Defines the style for body cells.
 *
 * When applied at class level, the style applies to all body cells.
 * When applied at property level, it is merged over the class-level style for that column:
 * properties set on the property win, and the others are inherited.
 *
 * If both enum color and hex color are specified, hex takes precedence.
 * This annotation has no number format option. For number/date formatting, use
 * `@Column(format = "...")`, which is applied as the body cells' number format.
 *
 * @param bold `true` makes the text bold. `false` (default) means "not specified": the value is
 *   inherited from lower-priority styles, so it cannot turn off bold from a theme or class-level
 *   annotation. Use a [ConditionalStyler] returning `CellStyle(bold = false)` for that.
 * @param italic `true` makes the text italic. `false` (default) means "not specified", as for [bold].
 * @param backgroundColor predefined background color
 * @param backgroundColorHex custom background color in "#RRGGBB" format (takes precedence)
 * @param fontColor predefined font color
 * @param fontColorHex custom font color in "#RRGGBB" format (takes precedence)
 * @param alignment text alignment within the cell
 * @param border border style for the cell
 */
@Target(AnnotationTarget.CLASS, AnnotationTarget.PROPERTY)
@Retention(AnnotationRetention.RUNTIME)
annotation class BodyStyle(
    val bold: Boolean = false,
    val italic: Boolean = false,
    val backgroundColor: StyleColor = StyleColor.NONE,
    val backgroundColorHex: String = "",
    val fontColor: StyleColor = StyleColor.NONE,
    val fontColorHex: String = "",
    val alignment: StyleAlignment = StyleAlignment.NONE,
    val border: StyleBorder = StyleBorder.NONE,
)

/**
 * Specifies a conditional style for a column based on cell values.
 *
 * The specified [ConditionalStyler] implementation will be instantiated and called
 * for each cell value to determine the style. The conditional style is merged with
 * other styles, taking highest precedence.
 *
 * Example:
 * ```kotlin
 * @Excel
 * data class Transaction(
 *     @Column("Amount")
 *     @ConditionalStyle(PriceStyler::class)
 *     val amount: Int,
 * )
 * ```
 *
 * @param styler the ConditionalStyler implementation class
 */
@Target(AnnotationTarget.PROPERTY)
@Retention(AnnotationRetention.RUNTIME)
annotation class ConditionalStyle(
    val styler: KClass<out ConditionalStyler<*>>,
)
