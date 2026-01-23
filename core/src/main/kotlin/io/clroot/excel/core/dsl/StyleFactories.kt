@file:Suppress("unused")

package io.clroot.excel.core.dsl

import io.clroot.excel.core.style.CellStyle
import io.clroot.excel.core.style.Color

/**
 * Creates a CellStyle with only fontColor set.
 * Convenience function for conditional styles.
 *
 * Example:
 * ```kotlin
 * column("Price", conditionalStyle = { value: Int? ->
 *     if (value != null && value < 0) fontColor(Color.RED) else null
 * }) { it.price }
 * ```
 *
 * @param color the font color to apply
 * @return a CellStyle with only fontColor set
 */
fun fontColor(color: Color): CellStyle = CellStyle(fontColor = color)

/**
 * Creates a CellStyle with only backgroundColor set.
 * Convenience function for conditional styles.
 *
 * Example:
 * ```kotlin
 * column("Status", conditionalStyle = { value: String? ->
 *     if (value == "WARNING") backgroundColor(Color.YELLOW) else null
 * }) { it.status }
 * ```
 *
 * @param color the background color to apply
 * @return a CellStyle with only backgroundColor set
 */
fun backgroundColor(color: Color): CellStyle = CellStyle(backgroundColor = color)
