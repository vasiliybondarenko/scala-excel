package poi.excel

import java.io.FileOutputStream
import java.time.Month

import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/12/15
 * Time: 4:22 PM 
 * Project: scala-excel
 */
object BasicReportWriter extends App{

  val path = "/Users/shredinger/Downloads/avail-report.xlsx"

  val month = Month.FEBRUARY
  val data = Map(
    (1, "OE-GVF") -> Array("s",	 "5.50",	 "3.50",	 "9.75",	 "CR",	 "CR",	 "6.00",	 "1.17",	 "CR"),
    (2, "OE-GVG") -> Array("4.25",	 "UMX",	 "UMX",	 "2.33",	 "4.33",	 "CR",	 "CR",	 "2.50",	 "AOG"),
    (3, "XX-XXX") -> Array("2015/01/12",	 "100.00",	 "UMX",	 "2.33",	 "4.33",	 "CR",	 "CR",	 "2.50",	 "AOG")
  )

  def matches(s:String, pattern:String):Boolean = s.matches(pattern)

  def getCellType(value:String) =  {
    if (matches(value, CellDataTypes.CURRENCY))  Currency(value) else
    if (matches(value, CellDataTypes.DATE))  DateCellType(value) else
    Text(value)
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
        val cellType = getCellType(value)

        val style = wb.createCellStyle()
        val format = wb.createDataFormat()
        style.setDataFormat(format.getFormat(cellType.formatString))
        cell.setCellStyle(style)
        cellType match {
          case _: Currency => cell.setCellValue(cellType.asInstanceOf[Currency].getValue)
          case _: DateCellType => cell.setCellValue(cellType.asInstanceOf[DateCellType].getValue)
          case _: Text => cell.setCellValue(cellType.asInstanceOf[Text].getValue)
        }

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
