package io.clroot.excel.render

import io.clroot.excel.core.model.ExcelDocument
import java.io.OutputStream

/**
 * Interface for rendering an ExcelDocument to various formats.
 */
interface ExcelRenderer {
    /**
     * Renders the document to the output stream.
     */
    fun render(document: ExcelDocument, output: OutputStream)
}
