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
 * @param format Format pattern for dates/numbers
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
 * When applied at property level, it overrides the class-level style for that column.
 *
 * If both enum color and hex color are specified, hex takes precedence.
 *
 * @param bold whether the text should be bold
 * @param italic whether the text should be italic
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
 * When applied at property level, it overrides the class-level style for that column.
 *
 * If both enum color and hex color are specified, hex takes precedence.
 * For number/date formatting, use @Column(format = "...") instead.
 *
 * @param bold whether the text should be bold
 * @param italic whether the text should be italic
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
