@file:Suppress("unused")

package io.clroot.excel.core.dsl

import io.clroot.excel.core.model.ColumnStyleConfig
import io.clroot.excel.core.model.ExcelDocument
import io.clroot.excel.core.model.Sheet
import io.clroot.excel.core.style.CellStyle

/**
 * Builder for constructing an [ExcelDocument].
 *
 * Use this builder via the [excel] DSL function to create Excel documents
 * with multiple sheets, styles, and data.
 *
 * Example:
 * ```kotlin
 * excel {
 *     sheet<User>("Users") {
 *         column("Name") { it.name }
 *         column("Age") { it.age }
 *         rows(users)
 *     }
 *     styles {
 *         header { bold(); backgroundColor(Color.LIGHT_BLUE) }
 *     }
 * }
 * ```
 *
 * @see excel
 * @see SheetBuilder
 */
@ExcelDslMarker
class ExcelBuilder(private val theme: ExcelTheme? = null) {
    private val sheets = mutableListOf<Sheet>()
    private var stylesConfig: StylesConfig? = null

    /**
     * Adds a new sheet to the Excel document.
     *
     * @param T the type of data objects that will populate this sheet
     * @param name the name of the sheet (displayed as tab name in Excel)
     * @param block the builder block to configure columns and rows
     */
    fun <T> sheet(
        name: String,
        block: SheetBuilder<T>.() -> Unit,
    ) {
        sheets.add(SheetBuilder<T>(name).apply(block).build())
    }

    /**
     * Configures global styles for the Excel document.
     *
     * Styles defined here apply to all sheets. They are merged on top of the theme passed to
     * [excel], and column-specific, inline and conditional styles are merged on top of them.
     *
     * @param block the builder block to configure header, body, and column-specific styles
     * @see StylesBuilder
     */
    fun styles(block: StylesBuilder.() -> Unit) {
        stylesConfig = StylesBuilder().apply(block).build()
    }

    internal fun build(): ExcelDocument {
        // Global styles are merged on top of the theme (theme < styles { header/body })
        val headerStyle = mergeNullable(theme?.headerStyle, stylesConfig?.headerStyle)
        val bodyStyle = mergeNullable(theme?.bodyStyle, stylesConfig?.bodyStyle)

        // Convert DSL ColumnStyleConfig to model ColumnStyleConfig
        val columnStyles =
            stylesConfig?.columnStyles?.mapValues { (_, config) ->
                ColumnStyleConfig(
                    headerStyle = config.headerStyle,
                    bodyStyle = config.bodyStyle,
                )
            } ?: emptyMap()

        return ExcelDocument(
            sheets = sheets.toList(),
            headerStyle = headerStyle,
            bodyStyle = bodyStyle,
            columnStyles = columnStyles,
        )
    }
}

private fun mergeNullable(
    lower: CellStyle?,
    higher: CellStyle?,
): CellStyle? =
    when {
        lower == null -> higher
        higher == null -> lower
        else -> lower.merge(higher)
    }
