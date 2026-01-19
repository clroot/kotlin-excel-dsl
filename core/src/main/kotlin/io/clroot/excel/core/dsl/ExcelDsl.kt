package io.clroot.excel.core.dsl

import io.clroot.excel.core.model.ExcelDocument
import io.clroot.excel.core.style.CellStyle

/**
 * DSL marker to prevent scope leakage.
 */
@DslMarker
annotation class ExcelDslMarker

/**
 * Theme interface for DSL.
 * Actual themes are defined in the theme module.
 */
interface ExcelTheme {
    val headerStyle: CellStyle
    val bodyStyle: CellStyle
}

/**
 * Entry point for building an Excel document using DSL.
 */
fun excel(block: ExcelBuilder.() -> Unit): ExcelDocument {
    return ExcelBuilder().apply(block).build()
}

/**
 * Entry point for building an Excel document with a theme.
 */
fun excel(
    theme: ExcelTheme,
    block: ExcelBuilder.() -> Unit,
): ExcelDocument {
    return ExcelBuilder(theme).apply(block).build()
}
