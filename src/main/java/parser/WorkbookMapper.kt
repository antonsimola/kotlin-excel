package parser

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.math.BigDecimal
import java.text.SimpleDateFormat
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.ZoneId
import java.util.*
import kotlin.reflect.KClass
import kotlin.reflect.KMutableProperty1
import kotlin.reflect.full.createInstance

class WorkbookMapper<T : Any>(
        val clazz: KClass<T>,
        val columnMappings: List<ColumnMapping<T>>? = null
) {
    private fun defaultMapping(): List<ColumnMapping<T>> {
        TODO()
    }

    fun read(workbook: Workbook): List<T> {
        val beans = mutableListOf<T>()
        val mapping = columnMappings ?: defaultMapping()
        workbook.iterator().forEach {
            val rowIterator = it.rowIterator()
            val headers = if (rowIterator.hasNext()) getHeaders(rowIterator.next()) else listOf() //TODO
            rowIterator.forEachRemaining {
                val bean: T = clazz.createInstance()
                for ((index, cell) in it.cellIterator().withIndex()) {
                    val header = headers[index]
                    val beanProperty = mapping.find { it.excelHeader == header }
                    beanProperty?.setPropertyFromCell?.invoke(bean, cell)
                }
                beans.add(bean)
            }
        }
        return beans
    }

    fun write(beans: List<T>): Workbook {
        if (columnMappings == null) TODO() //TODO
        val workbook = XSSFWorkbook()
        val sheet = workbook.createSheet()
        var index = 0
        createHeaderRow(sheet.createRow(index++), columnMappings.map { it.excelHeader }.toSet())
        for (bean in beans) {
            createBeanRow(sheet.createRow(index++), bean)
        }
        return workbook
    }

    private fun createHeaderRow(row: Row, headers: Set<String>) {
        headers.forEachIndexed { index, header -> row.createCell(index).setCellValue(header) }
    }

    private fun createBeanRow(row: Row, bean: T) {
        if (columnMappings == null) TODO()
        columnMappings.forEachIndexed { index, columnMapping ->
            val cell = row.createCell(index)
            columnMapping.setCellFromProperty.invoke(bean, cell)
        }
    }

    private fun getHeaders(row: Row): List<String> {
        return row.cellIterator().asSequence().map { getCellStringValue(it) }.toList()
    }
}

@JvmName("mapCellBigDecimal")
infix fun <T> KMutableProperty1<T, BigDecimal>.excelColumn(excelHeader: String): ColumnMapping<T> {
    val setPropertyFromCell = { bean: T, cell: Cell -> this.set(bean, BigDecimal(getCellNumberValue(cell))) }
    val setCellFromProperty = { bean: T, cell: Cell -> setCellValue(cell, this.get(bean)) }
    return ColumnMapping(excelHeader, setPropertyFromCell, setCellFromProperty)
}


@JvmName("mapCellString")
infix fun <T> KMutableProperty1<T, String>.excelColumn(excelHeader: String): ColumnMapping<T> {
    val setPropertyFromCell = { bean: T, cell: Cell -> this.set(bean, getCellStringValue(cell)) }
    val setCellFromProperty = { bean: T, cell: Cell -> setCellValue(cell, this.get(bean)) }
    return ColumnMapping(excelHeader, setPropertyFromCell, setCellFromProperty)
}

@JvmName("mapCellLong")
infix fun <T> KMutableProperty1<T, Long>.excelColumn(excelHeader: String): ColumnMapping<T> {
    val setPropertyFromCell = { bean: T, cell: Cell -> this.set(bean, parser.getCellNumberValue(cell).toLong()) }
    val setCellFromProperty = { bean: T, cell: Cell -> setCellValue(cell, this.get(bean)) }
    return ColumnMapping(excelHeader, setPropertyFromCell, setCellFromProperty)
}

@JvmName("mapCellDate")
infix fun <T> KMutableProperty1<T, Date?>.excelColumn(excelHeader: String): ColumnMapping<T> {
    val setPropertyFromCell = { bean: T, cell: Cell -> this.set(bean, parser.getCellDateValue(cell)) } //TODO
    val setCellFromProperty = { bean: T, cell: Cell -> setCellValue(cell, this.get(bean)) }
    return ColumnMapping(excelHeader, setPropertyFromCell, setCellFromProperty)
}

@JvmName("mapCellInt")
infix fun <T> KMutableProperty1<T, Int>.excelColumn(excelHeader: String): ColumnMapping<T> {
    val setPropertyFromCell = { bean: T, cell: Cell -> this.set(bean, parser.getCellNumberValue(cell).toInt()) }
    val setCellFromProperty = { bean: T, cell: Cell -> setCellValue(cell, this.get(bean)) }
    return ColumnMapping(excelHeader, setPropertyFromCell, setCellFromProperty)
}


private fun getCellStringValue(cell: Cell): String {
    return when (cell.cellTypeEnum) {
        null -> ""
        CellType.BLANK -> ""
        CellType._NONE -> ""
        CellType.NUMERIC -> TODO() //TODO
        CellType.STRING -> cell.stringCellValue ?: ""
        CellType.FORMULA -> TODO()
        CellType.BOOLEAN -> TODO()
        CellType.ERROR -> ""
    }
}

fun getCellNumberValue(cell: Cell): Double {
    return when (cell.cellTypeEnum) {
        null -> 0.0
        CellType.BLANK -> 0.0
        CellType._NONE -> 0.0
        CellType.NUMERIC -> cell.numericCellValue
        CellType.STRING -> cell.stringCellValue.toDouble()
        CellType.FORMULA -> TODO()
        CellType.BOOLEAN -> TODO()
        CellType.ERROR -> 0.0
    }
}

private fun getCellDateValue(cell: Cell): Date? {
    return when (cell.cellTypeEnum) {
        null -> null
        CellType.BLANK -> null
        CellType._NONE -> null
        CellType.NUMERIC -> cell.dateCellValue
        CellType.STRING -> SimpleDateFormat("yyyy-MM-dd").parse(cell.stringCellValue) //TODO
        CellType.FORMULA -> TODO()
        CellType.BOOLEAN -> TODO()
        CellType.ERROR -> null
    }
}

private fun setCellValue(cell: Cell, obj: Any?) {
    when (obj) {
        is String -> cell.setCellValue(obj)
        is Date -> cell.setCellValue(obj)
        is LocalDate -> cell.setCellValue(obj.toDate())
        is LocalDateTime -> cell.setCellValue(obj.toDate())
        is Boolean -> cell.setCellValue(obj)
        is Calendar -> cell.setCellValue(obj)
        is Number -> cell.setCellValue(obj.toDouble())
        else -> RuntimeException("unknown object type")
    }
}

fun LocalDate.toDate(): Date = Date.from(this.atStartOfDay(ZoneId.systemDefault()).toInstant())
fun LocalDateTime.toDate(): Date = Date.from(this.atZone(ZoneId.systemDefault()).toInstant())

class ColumnMapping<in T>(val excelHeader: String, val setPropertyFromCell: (T, Cell) -> Unit, val setCellFromProperty: (T, Cell) -> Unit)
