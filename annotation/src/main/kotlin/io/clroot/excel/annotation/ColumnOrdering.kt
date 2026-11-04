package io.clroot.excel.annotation

import kotlin.reflect.KClass
import kotlin.reflect.KProperty1
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.full.memberProperties
import kotlin.reflect.full.primaryConstructor
import kotlin.reflect.jvm.javaField

/**
 * Returns the [Column]-annotated properties of [klass] in column order.
 *
 * Both writing (`excelOf`) and parsing use this ordering:
 * 1. ascending [Column.order];
 * 2. for equal (or omitted) `order`, primary constructor parameter order;
 * 3. properties not declared in the primary constructor follow, in field declaration order.
 *
 * Kotlin reflection does not guarantee that `memberProperties` are returned in declaration order,
 * so this function must be used wherever column order matters.
 *
 * @param klass the `@Excel` annotated class
 * @return annotated properties sorted by column order; empty if none are annotated
 */
fun <T : Any> columnProperties(klass: KClass<T>): List<KProperty1<T, *>> {
    val constructorOrder =
        klass.primaryConstructor
            ?.parameters
            ?.mapNotNull { it.name }
            ?.withIndex()
            ?.associate { it.value to it.index }
            .orEmpty()
    val fieldOrder =
        klass.java.declaredFields
            .withIndex()
            .associate { it.value.name to it.index }

    return klass.memberProperties
        .filter { it.findAnnotation<Column>() != null }
        .sortedWith(
            compareBy<KProperty1<T, *>>(
                { it.findAnnotation<Column>()!!.order },
                { constructorOrder[it.name] ?: Int.MAX_VALUE },
                { it.javaField?.let { field -> fieldOrder[field.name] } ?: Int.MAX_VALUE },
            ).thenBy { it.name },
        )
}
