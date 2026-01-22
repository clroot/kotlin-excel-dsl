package io.clroot.excel.annotation

import io.clroot.excel.annotation.style.StyleConverter
import io.clroot.excel.core.ExcelConfigurationException
import io.clroot.excel.core.dsl.ExcelTheme
import io.clroot.excel.core.model.ColumnDefinition
import io.clroot.excel.core.model.ColumnWidth
import io.clroot.excel.core.model.ConditionalStyle
import io.clroot.excel.core.model.ExcelDocument
import io.clroot.excel.core.model.Sheet
import io.clroot.excel.core.style.CellStyle
import kotlin.reflect.KClass
import kotlin.reflect.KProperty1
import kotlin.reflect.full.createInstance
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.full.memberProperties

/**
 * Creates an ExcelDocument from annotated data class instances.
 */
inline fun <reified T : Any> excelOf(
    data: Iterable<T>,
    sheetName: String = "Sheet1",
    theme: ExcelTheme? = null,
): ExcelDocument {
    return excelOf(T::class, data, sheetName, theme)
}

/**
 * Creates an ExcelDocument from annotated data class instances.
 */
fun <T : Any> excelOf(
    klass: KClass<T>,
    data: Iterable<T>,
    sheetName: String = "Sheet1",
    theme: ExcelTheme? = null,
): ExcelDocument {
    val className = klass.qualifiedName ?: klass.simpleName ?: "Unknown"

    if (klass.findAnnotation<Excel>() == null) {
        throw ExcelConfigurationException(
            message = "Missing @Excel annotation",
            className = className,
            hint = "Add @Excel annotation to your data class: @Excel data class ${klass.simpleName}(...)",
        )
    }

    val columnProps =
        klass.memberProperties
            .filter { it.findAnnotation<Column>() != null }
            .sortedBy { it.findAnnotation<Column>()!!.order }

    if (columnProps.isEmpty()) {
        val allProps = klass.memberProperties.map { it.name }
        throw ExcelConfigurationException(
            message = "No properties annotated with @Column",
            className = className,
            hint = "Add @Column annotation to properties. Available properties: ${allProps.joinToString(", ")}",
        )
    }

    val typedProps = columnProps.filterIsInstance<KProperty1<T, *>>()

    // Extract class-level styles
    val classHeaderStyle = klass.findAnnotation<HeaderStyle>()?.let { StyleConverter.toCellStyle(it) }
    val classBodyStyle = klass.findAnnotation<BodyStyle>()?.let { StyleConverter.toCellStyle(it) }

    val columns =
        typedProps.map { prop ->
            val annotation = prop.findAnnotation<Column>()!!

            // Extract property-level styles
            val propHeaderStyle = prop.findAnnotation<HeaderStyle>()?.let { StyleConverter.toCellStyle(it) }
            val propBodyStyle = prop.findAnnotation<BodyStyle>()?.let { StyleConverter.toCellStyle(it) }

            // Merge styles: theme < class level < property level
            val mergedHeaderStyle = mergeStyles(theme?.headerStyle, classHeaderStyle, propHeaderStyle)
            val mergedBodyStyle = mergeStyles(theme?.bodyStyle, classBodyStyle, propBodyStyle)

            // Handle @ConditionalStyle
            val conditionalStyle =
                prop.findAnnotation<io.clroot.excel.annotation.ConditionalStyle>()?.let { ann ->
                    createConditionalStyle(ann.styler)
                }

            ColumnDefinition<T>(
                header = annotation.header,
                width = if (annotation.width > 0) ColumnWidth.Fixed(annotation.width) else ColumnWidth.Auto,
                format = annotation.format.ifEmpty { null },
                headerStyle = mergedHeaderStyle,
                bodyStyle = mergedBodyStyle,
                conditionalStyle = conditionalStyle,
                valueExtractor = { item: T -> prop.get(item) },
            )
        }

    return ExcelDocument(
        sheets =
            listOf(
                Sheet(
                    name = sheetName,
                    columns = columns,
                    dataSource = data,
                ),
            ),
    )
}

/**
 * Merges multiple styles in order. Later styles override earlier ones.
 *
 * @param styles styles to merge in order of precedence (first has lowest priority)
 * @return merged style, or null if all inputs are null
 */
private fun mergeStyles(vararg styles: CellStyle?): CellStyle? {
    return styles.filterNotNull().reduceOrNull { acc, style -> acc.merge(style) }
}

/**
 * Creates a ConditionalStyle from a ConditionalStyler class.
 *
 * @param stylerClass the ConditionalStyler implementation class
 * @return ConditionalStyle wrapping the styler
 */
@Suppress("UNCHECKED_CAST")
private fun createConditionalStyle(stylerClass: KClass<out ConditionalStyler<*>>): ConditionalStyle<Any?> {
    val styler = stylerClass.createInstance() as ConditionalStyler<Any?>
    return ConditionalStyle { value ->
        try {
            styler.style(value)
        } catch (e: ClassCastException) {
            throw ExcelConfigurationException(
                message = "ConditionalStyler type mismatch",
                className = stylerClass.qualifiedName ?: stylerClass.simpleName ?: "Unknown",
                hint =
                    "Ensure the ConditionalStyler generic type matches the property type. " +
                        "Got value of type: ${value?.let { it::class.simpleName } ?: "null"}",
            )
        }
    }
}
