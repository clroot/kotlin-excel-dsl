@file:Suppress("unused")

package io.clroot.excel

// Styles
typealias CellStyle = io.clroot.excel.core.style.CellStyle
typealias Color = io.clroot.excel.core.style.Color
typealias Alignment = io.clroot.excel.core.style.Alignment
typealias BorderStyle = io.clroot.excel.core.style.BorderStyle

// Column Style Config
typealias ColumnStyleConfig = io.clroot.excel.core.model.ColumnStyleConfig

// Exceptions
typealias ExcelException = io.clroot.excel.core.ExcelException

@Deprecated("Not thrown by the library. Catch ExcelWriteException or ExcelParseException instead.")
typealias ExcelDataException = io.clroot.excel.core.ExcelDataException
typealias ExcelWriteException = io.clroot.excel.core.ExcelWriteException
typealias ExcelConfigurationException = io.clroot.excel.core.ExcelConfigurationException

@Deprecated("Not thrown by the library. Missing columns are reported through ParseResult / ExcelParseException.")
typealias ColumnNotFoundException = io.clroot.excel.core.ColumnNotFoundException

@Deprecated("Not thrown by the library. Catch ExcelConfigurationException or ExcelWriteException instead.")
typealias StyleException = io.clroot.excel.core.StyleException
