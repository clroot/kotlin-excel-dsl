package io.clroot.excel.annotation

import io.clroot.excel.core.style.CellStyle

/**
 * Interface for defining conditional styles based on cell values.
 *
 * Implementing classes must have a no-arg constructor as they will be
 * instantiated via reflection at runtime.
 *
 * Example:
 * ```kotlin
 * class PriceStyler : ConditionalStyler<Int> {
 *     override fun style(value: Int?): CellStyle? = when {
 *         value == null -> null
 *         value < 0 -> CellStyle(fontColor = Color.RED)
 *         value > 1_000_000 -> CellStyle(fontColor = Color.GREEN, bold = true)
 *         else -> null
 *     }
 * }
 * ```
 *
 * @param T the type of the cell value
 */
interface ConditionalStyler<T> {
    /**
     * Returns the style to apply for the given cell value.
     *
     * @param value the cell value, may be null
     * @return the style to apply, or null if no conditional style should be applied
     */
    fun style(value: T?): CellStyle?
}
