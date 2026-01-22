@file:Suppress("unused")

package io.clroot.excel.render

import io.clroot.excel.core.model.ExcelDocument
import java.io.ByteArrayOutputStream
import java.io.OutputStream

/**
 * Extension function to write an ExcelDocument to an OutputStream.
 */
fun ExcelDocument.writeTo(
    output: OutputStream,
    renderer: ExcelRenderer = PoiRenderer(),
) {
    renderer.render(this, output)
}

/**
 * Extension function to render an ExcelDocument to a ByteArray.
 */
fun ExcelDocument.toByteArray(renderer: ExcelRenderer = PoiRenderer()): ByteArray =
    ByteArrayOutputStream().use { output ->
        renderer.render(this, output)
        output.toByteArray()
    }
