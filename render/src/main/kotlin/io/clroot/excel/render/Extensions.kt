@file:Suppress("unused")

package io.clroot.excel.render

import io.clroot.excel.core.model.ExcelDocument
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
