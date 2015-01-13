package poi.excel.api

import java.io.FileOutputStream

import org.apache.poi.xssf.usermodel.{XSSFCell, XSSFWorkbook}
import poi.excel._
import poi.excel.api.ExcelReporter.{ReportFile, ReportData, ReportInfo, Reporter}

object SimpleReporter extends Reporter{

  val sheetName = "Avail report - All prod"

  val columnWidth = 20

  val destinationPath = "/Users/shredinger/Downloads/api_test.xlsx"


  override def run(reportInfo: ReportInfo, reportData: ReportData): ReportFile = {
    val wb = new XSSFWorkbook()
    val sheet = wb.createSheet(sheetName)
    val headerRow = sheet.createRow(0)
    val columnsToShow = reportInfo.colShowList.toSet
    val colMetaData = reportInfo.colInfo.toArray

    sheet.setDefaultColumnWidth(columnWidth)

    //header
    var columnNumber = 1
    for(item <- reportInfo.colInfo if columnsToShow.contains(item.field)) {
      val cell = headerRow.createCell(columnNumber)
      cell.setCellValue(item.title)

      columnNumber = columnNumber + 1
    }

    //body
    var rowNumber = 1
    for(rowItem <- reportData){
      val row = sheet.createRow(rowNumber + 1)
      row.createCell(0).setCellValue(rowItem.head)

      var columnNumber = 0
      for(cellValue <- rowItem){
        val cell:XSSFCell = row.createCell(columnNumber)
        val value = cellValue

        val dataConverter = CellDataConverter(colMetaData(columnNumber).dataType)

        val style = wb.createCellStyle()
        val format = wb.createDataFormat()
        style.setDataFormat(format.getFormat(dataConverter.getDataFormat))
        cell.setCellStyle(style)
        dataConverter.assignValue(value, cell)

        columnNumber = columnNumber + 1
      }

      rowNumber = rowNumber + 1
    }

    val resultFile = new FileOutputStream(destinationPath)
    wb.write(resultFile)
    resultFile.close

  }
}