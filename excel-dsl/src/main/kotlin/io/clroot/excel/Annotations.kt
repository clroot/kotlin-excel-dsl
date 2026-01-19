@file:Suppress("unused")

package io.clroot.excel

import io.clroot.excel.core.dsl.ExcelTheme
import io.clroot.excel.core.model.ExcelDocument

/**
 * Marks a class as an Excel-exportable data class.
 */
typealias Excel = io.clroot.excel.annotation.Excel

/**
 * Configures a property as an Excel column.
 */
typealias Column = io.clroot.excel.annotation.Column

/**
 * Creates an ExcelDocument from annotated data class instances.
 */
inline fun <reified T : Any> excelOf(
    data: Iterable<T>,
    sheetName: String = "Sheet1",
    theme: ExcelTheme? = null,
): ExcelDocument =
    io.clroot.excel.annotation
        .excelOf(data, sheetName, theme)
