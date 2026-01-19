package io.clroot.excel.annotation

import io.clroot.excel.core.ExcelConfigurationException
import io.clroot.excel.core.dsl.ExcelTheme
import io.clroot.excel.core.model.*
import kotlin.reflect.KClass
import kotlin.reflect.KProperty1
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

    val columns =
        typedProps.map { prop ->
            val annotation = prop.findAnnotation<Column>()!!
            ColumnDefinition<T>(
                header = annotation.header,
                width = if (annotation.width > 0) ColumnWidth.Fixed(annotation.width) else ColumnWidth.Auto,
                format = annotation.format.ifEmpty { null },
                valueExtractor = { item: T -> prop.get(item) },
            )
        }

    val rows =
        data.map { item ->
            Row(
                cells =
                    typedProps.map { prop ->
                        Cell(value = prop.get(item))
                    },
            )
        }

    return ExcelDocument(
        sheets =
            listOf(
                Sheet(
                    name = sheetName,
                    columns = columns,
                    rows = rows,
                ),
            ),
        headerStyle = theme?.headerStyle,
        bodyStyle = theme?.bodyStyle,
    )
}
