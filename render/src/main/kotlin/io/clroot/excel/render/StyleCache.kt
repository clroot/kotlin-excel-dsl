package io.clroot.excel.render

import io.clroot.excel.core.style.Alignment
import io.clroot.excel.core.style.BorderStyle
import io.clroot.excel.core.style.CellStyle
import io.clroot.excel.core.style.Color
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.ss.usermodel.BorderStyle as PoiBorderStyle
import org.apache.poi.ss.usermodel.CellStyle as PoiCellStyle

/**
 * Caches POI CellStyle objects to prevent duplicate creation.
 *
 * POI workbooks have a limit of approximately 64,000 unique cell styles.
 * This cache ensures that identical [CellStyle] configurations are reused,
 * significantly reducing the number of POI styles created.
 *
 * @property workbook the POI workbook to create styles for
 */
internal class StyleCache(private val workbook: Workbook) {
    private val cache = mutableMapOf<CellStyle, PoiCellStyle>()
    private val fontCache = mutableMapOf<FontKey, Font>()

    /**
     * Gets or creates a POI CellStyle for the given domain CellStyle.
     *
     * If an identical style already exists in the cache, it is returned.
     * Otherwise, a new POI style is created, cached, and returned.
     *
     * @param style the domain CellStyle to convert
     * @return the corresponding POI CellStyle
     */
    fun getOrCreate(style: CellStyle): PoiCellStyle {
        return cache.getOrPut(style) { createPoiStyle(style) }
    }

    private fun createPoiStyle(style: CellStyle): PoiCellStyle {
        val poiStyle = workbook.createCellStyle()

        // Background color
        style.backgroundColor?.let { color ->
            poiStyle.fillForegroundColor = createXSSFColor(color).indexed
            poiStyle.setFillForegroundColor(createXSSFColor(color))
            poiStyle.fillPattern = FillPatternType.SOLID_FOREGROUND
        }

        // Font (cached separately)
        val fontKey =
            FontKey(
                bold = style.bold,
                italic = style.italic,
                fontColor = style.fontColor,
            )
        val font = fontCache.getOrPut(fontKey) { createFont(fontKey) }
        poiStyle.setFont(font)

        // Alignment
        style.alignment?.let { alignment ->
            poiStyle.alignment =
                when (alignment) {
                    Alignment.LEFT -> HorizontalAlignment.LEFT
                    Alignment.CENTER -> HorizontalAlignment.CENTER
                    Alignment.RIGHT -> HorizontalAlignment.RIGHT
                }
        }

        // Border
        style.border?.let { border ->
            val poiBorder =
                when (border) {
                    BorderStyle.NONE -> PoiBorderStyle.NONE
                    BorderStyle.THIN -> PoiBorderStyle.THIN
                    BorderStyle.MEDIUM -> PoiBorderStyle.MEDIUM
                    BorderStyle.THICK -> PoiBorderStyle.THICK
                }
            poiStyle.borderTop = poiBorder
            poiStyle.borderBottom = poiBorder
            poiStyle.borderLeft = poiBorder
            poiStyle.borderRight = poiBorder
        }

        // Number format
        style.numberFormat?.let { format ->
            poiStyle.dataFormat = workbook.createDataFormat().getFormat(format)
        }

        return poiStyle
    }

    private fun createFont(fontKey: FontKey): Font {
        val font = workbook.createFont()
        if (fontKey.bold) {
            font.bold = true
        }
        if (fontKey.italic) {
            font.italic = true
        }
        fontKey.fontColor?.let { color ->
            if (font is XSSFFont) {
                font.setColor(createXSSFColor(color))
            }
        }
        return font
    }

    private fun createXSSFColor(color: Color): XSSFColor {
        return XSSFColor(byteArrayOf(color.red.toByte(), color.green.toByte(), color.blue.toByte()), null)
    }

    /**
     * Key for font caching.
     */
    private data class FontKey(
        val bold: Boolean,
        val italic: Boolean,
        val fontColor: Color?,
    )
}
