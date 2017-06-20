package test

import model.Person
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.junit.Test
import parser.WorkbookMapper
import parser.excelColumn
import parser.toDate
import java.io.File
import java.io.FileOutputStream
import java.math.BigDecimal
import java.time.LocalDate
import kotlin.test.assertEquals


class ImportTest {
    @Test
    fun testExplicitMapping() {
        val testXls = javaClass.classLoader.getResourceAsStream("test/test.xlsx")
        val wb = WorkbookFactory.create(testXls)
        val excelHeadersToProps = listOf(
                Person::name excelColumn "name",
                Person::wage excelColumn "wage",
                Person::birthDayOld excelColumn "birthDayOld",
                Person::age excelColumn "age"
        )
        val parser = WorkbookMapper<Person>(Person::class, excelHeadersToProps)
        val people: List<Person> = parser.read(wb)
        val expected = listOf(
                Person(name = "Anton", age = 23, wage = BigDecimal(4287800), birthDayOld = LocalDate.of(1994, 10, 31).toDate()),
                Person(name = "Ilari", age = 22, wage = BigDecimal(50.50), birthDayOld = LocalDate.of(1994, 10, 31).toDate()),
                Person(name = "Juho", age = 0, wage = BigDecimal(0))
        )
        assertEquals(expected[0], people[0])
        assertEquals(expected[1], people[1])
        assertEquals(expected[2], people[2])


        // assertEquals(expected, people)
        people.forEach { println(it) }
    }

    @Test
    fun testImplicitMapping() {
        val testExcel = javaClass.classLoader.getResourceAsStream("test/test.xlsx")
        val wb = WorkbookFactory.create(testExcel)
        val parser = WorkbookMapper(Person::class)
        val beans: List<Person> = parser.read(wb)
        //TODO
        val expected = listOf(
                Person(name = "Anton", age = 23, wage = BigDecimal(4287800), height = 4.5f, weight = 3.5),
                Person(name = "Ilari", age = 22, wage = BigDecimal(50.50), height = 5.5f, weight = 3.2),
                Person(name = "Juho", age = 0, wage = BigDecimal(0), height = 0.0f, weight = 0.0)
        )
        assertEquals(expected, beans)
        beans.forEach { println(it) }
    }

    @Test
    fun importAndExport() {
        val propsToHeaders = listOf(
                Person::name excelColumn "Full name",
                Person::wage excelColumn "The wage",
                Person::birthDayOld excelColumn "Birthday",
                Person::age excelColumn "Age"
        )
        val parser = WorkbookMapper(Person::class, propsToHeaders)
        val people = listOf(
                Person(name = "Anton", age = 23, wage = BigDecimal(4287800), birthDayOld = LocalDate.of(1994, 10, 31).toDate()),
                Person(name = "Ilari", age = 22, wage = BigDecimal(50.50), birthDayOld = LocalDate.of(1994, 10, 31).toDate()),
                Person(name = "Juho", age = 0, wage = BigDecimal(0))
        )
        val wb = parser.write(people)
        val testExcel = File.createTempFile("temp", ".xlsx")
        wb.write(FileOutputStream(testExcel))
        val importedPeople = parser.read(WorkbookFactory.create(testExcel))

        assertEquals(people[0], importedPeople[0])
        assertEquals(people[1], importedPeople[1])
        assertEquals(people[2], importedPeople[2])

        assertEquals(importedPeople, people)
    }

}
