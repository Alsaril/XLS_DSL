import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.util.RegionUtil


class Document {

    private val xssfWorkbook = XSSFWorkbook()

    fun sheet(name: String, rows: Sheet.() -> Unit) {
        val sheet = Sheet(name, xssfWorkbook)
        rows.invoke(sheet)
    }

    fun save(name: String) {
        val file = FileOutputStream(name)
        xssfWorkbook.write(file)
        file.flush()
        file.close()
    }
}

class Sheet(name: String, private val wb: XSSFWorkbook) {

    private val xssfSheet = wb.createSheet(name)
    private var id = 0

    fun row(cells: Row.() -> Unit) {
        val row = Row(xssfSheet, id++, wb)
        cells.invoke(row)
    }

    @Deprecated("Wrong scope", level = DeprecationLevel.ERROR)
    fun document(file: String, sheets: Document.() -> Unit) {
    }
}

class Row(private val sheet: XSSFSheet, private val rowId: Int, wb: XSSFWorkbook) {

    private val xssfRow = sheet.createRow(rowId)
    private var id = 0
    private val style = wb.createCellStyle()

    init {
        val borderStyle = BorderStyle.DOUBLE
        val borderColor = IndexedColors.BLACK.getIndex()
        style.setBorderBottom(borderStyle)
        style.setBottomBorderColor(borderColor)
        style.setBorderLeft(borderStyle)
        style.setLeftBorderColor(borderColor)
        style.setBorderRight(borderStyle)
        style.setRightBorderColor(borderColor)
        style.setBorderTop(borderStyle)
        style.setTopBorderColor(borderColor)
    }

    fun cell(value: String, width: Int = 1) {
        assert(width >= 1)
        val cell = xssfRow.createCell(id)
        cell.setCellValue(value)
        cell.setCellStyle(style)
        if (width > 1) {
            sheet.addMergedRegion(CellRangeAddress(
                    rowId,
                    rowId,
                    id,
                    id + width
            ))
            RegionUtil.setBorderRight(BorderStyle.DOUBLE, CellRangeAddress(
                    rowId,
                    rowId,
                    id,
                    id + width
            ), sheet)
        }
        id += width
    }

    @Deprecated("Wrong scope", level = DeprecationLevel.ERROR)
    fun sheet(name: String, rows: Sheet.() -> Unit) {
    }
}

fun document(file: String, sheets: Document.() -> Unit) {
    val document = Document()
    sheets.invoke(document)
    document.save(file)
}

fun main(args: Array<String>) {
    document("out.xlsx") {
        sheet("sheet1") {
            row {
                cell("cell1")
                cell("cell2")
                cell("cell3", width = 10)
            }

            row {
                cell("cell1")
                cell("cell2")
                cell("cell3", width = 10)
            }
        }
    }
}