package poi.excel

import java.awt.Color
import java.io.FileOutputStream
import java.time.Month

import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.{XSSFColor, XSSFWorkbook}

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/12/15
 * Time: 4:22 PM 
 * Project: scala-excel
 */
object BasicReportWriter extends App{

  val path = "/Users/shredinger/Downloads/avail-report.xls"

  val month = Month.FEBRUARY
  val data = Map(
    (1, "OE-GVF") -> Array("s",	 "5.50",	 "3.50",	 "9.75",	 "CR",	 "CR",	 "6.00",	 "1.17",	 "CR"),
    (2, "OE-GVG") -> Array("4.25",	 "UMX",	 "UMX",	 "2.33",	 "4.33",	 "CR",	 "CR",	 "2.50",	 "AOG")
  )

  def getStyle(value:String, wb:XSSFWorkbook) = {
    val style = wb.createCellStyle
    style.setFillBackgroundColor(new XSSFColor(Color.orange))

    val format = wb.createDataFormat()
    style.setDataFormat(format.getFormat("#,##0.0000"))

    style
  }

  def sheetName = "Avail report - All prod"

  def days = Array[String]("Sat", "Sun", "Mon", "Tue", "Wed", "Tru", "Fri", "Sat", "Sun")

  def generate(data:Map[(Int, String), Array[String]]) = {
    val wb = new XSSFWorkbook()
    val sheet = wb.createSheet(sheetName)
    
    //header
    val headerRow = sheet.createRow(0)
    headerRow.createCell(0).setCellValue("")
    sheet.createRow(1).createCell(0).setCellValue("")
    for(idx <- 1 to days.size) {
      headerRow.createCell(idx).setCellValue(days(idx - 1))
    }

    //body
    for(key @ (rowNumber, rowName) <- data.keys) {
      val row = sheet.createRow(rowNumber + 1)
      row.createCell(0).setCellValue(rowName)
      for(idx <- 1 to data(key).size) {
        val cell = row.createCell(idx)
        val value = data(key)(idx - 1)
        cell.setCellValue(value)
        cell.setCellStyle(getStyle(value, wb))


      }
    }

    //merging cells
    sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0))


    //style
    val footerStyle = wb.createCellStyle
    val footerFont = wb.createFont
    footerFont.setBold(true)
    footerStyle.setFont(footerFont)

    //footer
    val footerRow = sheet.createRow(data.size + 2)
    val firstCell = footerRow.createCell(0)
    firstCell.setCellValue("Total Block Time: ")
    firstCell.setCellStyle(footerStyle)
    for(idx <- 1 to days.size) {
      val cell = footerRow.createCell(idx)
      val range = new CellRangeAddress(1, data.size, idx, idx)
      cell.setCellFormula(s"SUM(${range.formatAsString})")
      cell.setCellStyle(footerStyle)
    }

    //save
    val resultFile = new FileOutputStream(path)
    wb.write(resultFile)
    resultFile.close
    
  }





  generate(data)
}
