package io.clroot.excel.core.model

/**
 * Represents an Excel formula.
 *
 * Formulas are evaluated by Excel when the file is opened.
 * The expression can optionally include or omit the leading '=' sign.
 *
 * Example:
 * ```kotlin
 * column("합계") { formula("SUM(A2:A100)") }
 * column("평균") { formula("=AVERAGE(B2:B100)") }  // Leading '=' is optional
 * ```
 *
 * @property expression the original formula expression as provided
 * @throws IllegalArgumentException if the expression is blank after removing '='
 */
data class Formula(val expression: String) {
    /**
     * Returns the expression without the leading '=' if present, trimmed of whitespace.
     * This is the format expected by Apache POI's setCellFormula() method.
     */
    val normalizedExpression: String
        get() = expression.removePrefix("=").trim()

    init {
        require(normalizedExpression.isNotBlank()) {
            "Formula expression cannot be blank"
        }
    }
}

/**
 * Creates a Formula with the given expression.
 *
 * This is a convenience function for creating Formula objects in column definitions.
 *
 * Example:
 * ```kotlin
 * cell(formula("SUM(A1:A10)"))
 * cell(formula("=AVERAGE(B2:B100)"))  // Leading '=' is optional
 * ```
 *
 * @param expression the formula expression
 * @return a Formula object
 */
fun formula(expression: String): Formula = Formula(expression)
