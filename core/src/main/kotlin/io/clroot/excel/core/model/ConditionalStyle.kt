package io.clroot.excel.core.model

import io.clroot.excel.core.style.CellStyle

/**
 * Defines a conditional style that applies based on cell value.
 *
 * The style function receives the cell value and returns a [CellStyle] if the condition
 * is met, or null if no conditional style should be applied.
 *
 * When a conditional style returns a non-null [CellStyle], it is merged with the column's
 * base body style (if any), with the conditional style taking precedence.
 *
 * Example:
 * ```kotlin
 * column("Amount") {
 *     conditionalStyle { value: Int ->
 *         when {
 *             value < 0 -> CellStyle(fontColor = Color.RED)
 *             value > 1000000 -> CellStyle(fontColor = Color.GREEN, bold = true)
 *             else -> null
 *         }
 *     }
 * }
 * ```
 *
 * @param T the type of the cell value
 * @property styleFunction the function that determines the style based on value
 */
data class ConditionalStyle<T>(
    val styleFunction: (T) -> CellStyle?,
) {
    /**
     * Evaluates the conditional style for the given value.
     *
     * @param value the cell value to evaluate
     * @return the style to apply, or null if no conditional style matches
     */
    fun evaluate(value: T): CellStyle? = styleFunction(value)
}
