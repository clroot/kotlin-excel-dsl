package io.clroot.excel.render

import io.clroot.excel.core.ExcelWriteException
import io.clroot.excel.core.model.ExcelDocument
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import java.io.OutputStream

/**
 * Renders ExcelDocument to .xlsx format using Apache POI (SXSSF for streaming).
 *
 * This renderer uses Apache POI's SXSSF (Streaming Usermodel API) for memory-efficient
 * handling of large datasets. It includes style caching to avoid exceeding POI's
 * style limit (approximately 64,000 styles per workbook).
 *
 * @property rowAccessWindowSize the number of rows to keep in memory (default: 100)
 */
class PoiRenderer(
    private val rowAccessWindowSize: Int = 100,
) : ExcelRenderer {
    override fun render(
        document: ExcelDocument,
        output: OutputStream,
    ) {
        try {
            SXSSFWorkbook(rowAccessWindowSize).use { workbook ->
                val styleCache = StyleCache(workbook)
                val styleResolver = StyleResolver.from(document)

                document.sheets.forEach { sheetModel ->
                    val sheet = workbook.createSheet(sheetModel.name)
                    val sheetRenderer =
                        SheetRenderer(
                            sheet = sheet,
                            sheetModel = sheetModel,
                            styleResolver = styleResolver,
                            styleCache = styleCache,
                        )
                    sheetRenderer.render()
                }

                workbook.write(output)
            }
        } catch (e: Exception) {
            throw ExcelWriteException(
                message = "Failed to write Excel document: ${e.message}",
                cause = e,
            )
        }
    }
}
