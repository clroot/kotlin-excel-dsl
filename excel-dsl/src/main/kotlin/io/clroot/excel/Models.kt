@file:Suppress("unused")

package io.clroot.excel

import io.clroot.excel.core.model.ColumnWidth as CoreColumnWidth

// Models
typealias ExcelDocument = io.clroot.excel.core.model.ExcelDocument
typealias Sheet = io.clroot.excel.core.model.Sheet
typealias ColumnDefinition<T> = io.clroot.excel.core.model.ColumnDefinition<T>
typealias ColumnWidth = io.clroot.excel.core.model.ColumnWidth
typealias HeaderGroup = io.clroot.excel.core.model.HeaderGroup

// Column width extensions
val Int.chars: ColumnWidth get() = CoreColumnWidth.Fixed(this)
val Int.percent: ColumnWidth get() = CoreColumnWidth.Percent(this)
val auto: ColumnWidth = CoreColumnWidth.Auto
