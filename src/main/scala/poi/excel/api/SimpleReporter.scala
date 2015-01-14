package poi.excel.api

import java.io.FileOutputStream

import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.{XSSFCell, XSSFWorkbook}
import poi.excel._
import poi.excel.api.ExcelReporter._


object SimpleReporter extends Reporter{

  val sheetName = "Avail report - All prod"

  val columnWidth = 20

  val destinationPath = "/Users/shredinger/Downloads/api_test.xlsx"


  override def run(reportInfo: ReportInfo, reportData: ReportData): ReportFile = {
    val wb = new XSSFWorkbook()
    val sheet = wb.createSheet(sheetName)
    val headerRow = sheet.createRow(0)
    val colMetaData = getFilteredColumns(reportInfo).toArray

    sheet.setDefaultColumnWidth(columnWidth)

    //header
    var columnNumber = 0
    for(item <- getFilteredColumns(reportInfo)) {
      val cell = headerRow.createCell(columnNumber)
      cell.setCellValue(item.title)
      columnNumber = columnNumber + 1
    }

    //body
    var rowNumber = 1
    for(rowItem <- prepareData(reportInfo, reportData)){
      val row = sheet.createRow(rowNumber + 1)
      row.createCell(0).setCellValue(rowItem.head)

      var columnNumber = 0
      val cells = getVisibleCells(rowItem, reportInfo)
      for(cellValue <- cells){
        val cell:XSSFCell = row.createCell(columnNumber)
        val value = cellValue

        val dataConverter = CellDataConverter(colMetaData(columnNumber).dataType)
        val columnFormatString = colMetaData(columnNumber).format

        val style = wb.createCellStyle()
        val format = wb.createDataFormat()
        style.setDataFormat(format.getFormat(
          columnFormatString match {
            case Some(value) => value
            case _ => dataConverter.getDataFormat
          }
        ))
        cell.setCellStyle(style)
        dataConverter.assignValue(value, cell)

        columnNumber = columnNumber + 1
      }

      rowNumber = rowNumber + 1
    }

    //grouping
    //for test
    sheet.addMergedRegion(new CellRangeAddress(2, 3, 0, 0))

    val resultFile = new FileOutputStream(destinationPath)
    wb.write(resultFile)
    resultFile.close

  }

  def prepareData(reportInfo: ReportInfo, reportData: ReportData):ReportData = reportInfo.rowGroup match {
    case List(_) => DataSorter.sort(reportInfo, reportData)
    case Nil => reportData
  }

  def getFilteredColumns(reportInfo: ReportInfo): List[ColInfo] = {
    val columnsToShow = reportInfo.colShowList.toSet
    reportInfo.colInfo.filter(c => columnsToShow.contains(c.field))
  }

  def getVisibleCells(row: List[String], reportInfo: ReportInfo): List[String] = {
    val columnsToShow = reportInfo.colShowList.toSet
    (reportInfo.colInfo zip row)
      .filter(pair => columnsToShow.contains(pair._1.field))
      .map(pair => pair._2)
  }
}