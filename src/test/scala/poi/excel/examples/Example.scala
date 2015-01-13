package poi.excel.examples

import java.io.FileOutputStream

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.scalatest.{FlatSpec, Matchers}

class Example extends FlatSpec with Matchers{
  val file = "/Users/shredinger/Downloads/MGMT Availability Tool 1.0.xls"

  it should "read xls file " in {
    val wb = new XSSFWorkbook()
    val sheet = wb.createSheet("Test")

    val headerRow = sheet.createRow(0)
    headerRow.createCell(0).setCellValue("February")
    headerRow.createCell(1).setCellValue("A1")
    headerRow.createCell(2).setCellValue("A2")
    headerRow.createCell(3).setCellValue("A3")

//    val row = sheet.createRow(1)
//    row.createCell(0).setCellValue("Row 1")
//    row.createCell(1).setCellValue("value 1")
//    row.createCell(1).setCellValue("value 2")
//    row.createCell(1).setCellValue("value 3")

    val resultFile = new FileOutputStream("/Users/shredinger/Downloads/simplest_report.xls")
    wb.write(resultFile)
    resultFile.close
  }
}
