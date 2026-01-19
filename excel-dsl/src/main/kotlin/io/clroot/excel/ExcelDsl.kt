@file:Suppress("unused")

package io.clroot.excel

import io.clroot.excel.core.dsl.ExcelBuilder
import io.clroot.excel.core.dsl.ExcelTheme
import io.clroot.excel.core.model.ExcelDocument
import io.clroot.excel.render.writeTo as renderWriteTo
import java.io.OutputStream

/**
 * Entry point for building an Excel document using DSL.
 */
fun excel(block: ExcelBuilder.() -> Unit): ExcelDocument =
    io.clroot.excel.core.dsl.excel(block)

/**
 * Entry point for building an Excel document with a theme.
 */
fun excel(theme: ExcelTheme, block: ExcelBuilder.() -> Unit): ExcelDocument =
    io.clroot.excel.core.dsl.excel(theme, block)

/**
 * Extension function to write an ExcelDocument to an OutputStream.
 */
fun ExcelDocument.writeTo(output: OutputStream) = renderWriteTo(output)
