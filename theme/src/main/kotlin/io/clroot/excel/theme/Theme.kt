package io.clroot.excel.theme

import io.clroot.excel.core.dsl.ExcelTheme
import io.clroot.excel.core.style.*

/**
 * Defines a theme for Excel styling.
 */
interface Theme : ExcelTheme {
    override val headerStyle: CellStyle
    override val bodyStyle: CellStyle

    companion object {
        val Modern: Theme = ModernTheme
        val Minimal: Theme = MinimalTheme
        val Classic: Theme = ClassicTheme
    }
}

/**
 * Modern theme with bold header and blue background.
 */
object ModernTheme : Theme {
    override val headerStyle = CellStyle(
        backgroundColor = Color(59, 89, 152),
        fontColor = Color.WHITE,
        bold = true,
        alignment = Alignment.CENTER,
        border = BorderStyle.THIN
    )

    override val bodyStyle = CellStyle(
        border = BorderStyle.THIN
    )
}

/**
 * Minimal theme with subtle styling.
 */
object MinimalTheme : Theme {
    override val headerStyle = CellStyle(
        bold = true,
        alignment = Alignment.LEFT,
        border = BorderStyle.THIN
    )

    override val bodyStyle = CellStyle(
        border = BorderStyle.NONE
    )
}

/**
 * Classic theme with traditional Excel look.
 */
object ClassicTheme : Theme {
    override val headerStyle = CellStyle(
        backgroundColor = Color.LIGHT_GRAY,
        fontColor = Color.BLACK,
        bold = true,
        alignment = Alignment.CENTER,
        border = BorderStyle.MEDIUM
    )

    override val bodyStyle = CellStyle(
        border = BorderStyle.THIN
    )
}
