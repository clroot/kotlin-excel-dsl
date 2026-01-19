package io.clroot.excel.annotation

/**
 * Marks a class as an Excel-exportable data class.
 */
@Target(AnnotationTarget.CLASS)
@Retention(AnnotationRetention.RUNTIME)
annotation class Excel

/**
 * Configures a property as an Excel column.
 *
 * @param header The column header text
 * @param width Column width in characters (0 = auto)
 * @param format Format pattern for dates/numbers
 * @param order Column order (lower values appear first)
 */
@Target(AnnotationTarget.PROPERTY)
@Retention(AnnotationRetention.RUNTIME)
annotation class Column(
    val header: String,
    val width: Int = 0,
    val format: String = "",
    val order: Int = Int.MAX_VALUE,
)
