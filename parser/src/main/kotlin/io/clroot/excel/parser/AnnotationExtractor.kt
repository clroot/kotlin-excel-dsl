package io.clroot.excel.parser

import io.clroot.excel.annotation.Column
import io.clroot.excel.annotation.Excel
import io.clroot.excel.core.ExcelConfigurationException
import kotlin.reflect.KClass
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.full.memberProperties
import kotlin.reflect.jvm.jvmErasure

/**
 * Extracts column metadata from @Excel/@Column annotated classes.
 */
object AnnotationExtractor {
    /**
     * Extract column metadata from the given class.
     */
    fun <T : Any> extractColumns(klass: KClass<T>): List<ColumnMeta> {
        val className = klass.qualifiedName ?: klass.simpleName ?: "Unknown"

        if (klass.findAnnotation<Excel>() == null) {
            throw ExcelConfigurationException(
                message = "Missing @Excel annotation",
                className = className,
                hint = "Add @Excel annotation to your data class: @Excel data class ${klass.simpleName}(...)",
            )
        }

        // Get constructor parameter order for stable sorting when order values are equal
        val constructorParamOrder =
            getConstructorParameterOrder(klass)
                .withIndex()
                .associate { it.value to it.index }

        val columnProps =
            klass.memberProperties
                .filter { it.findAnnotation<Column>() != null }
                .sortedWith(
                    compareBy(
                        { it.findAnnotation<Column>()!!.order },
                        { constructorParamOrder[it.name] ?: Int.MAX_VALUE },
                    ),
                )

        if (columnProps.isEmpty()) {
            val allProps = klass.memberProperties.map { it.name }
            throw ExcelConfigurationException(
                message = "No properties annotated with @Column",
                className = className,
                hint = "Add @Column annotation to properties. Available properties: ${allProps.joinToString(", ")}",
            )
        }

        return columnProps.map { prop ->
            val annotation = prop.findAnnotation<Column>()!!
            ColumnMeta(
                header = annotation.header,
                aliases = annotation.aliases,
                propertyName = prop.name,
                propertyType = prop.returnType.jvmErasure,
                isNullable = prop.returnType.isMarkedNullable,
            )
        }
    }

    /**
     * Get the primary constructor parameter order for the given class.
     */
    fun <T : Any> getConstructorParameterOrder(klass: KClass<T>): List<String> {
        val constructor =
            klass.constructors.firstOrNull()
                ?: throw ExcelConfigurationException(
                    message = "No constructor found",
                    className = klass.qualifiedName,
                    hint = "Ensure the class has a primary constructor",
                )

        return constructor.parameters.mapNotNull { it.name }
    }
}
