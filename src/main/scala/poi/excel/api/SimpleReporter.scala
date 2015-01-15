package poi.excel.api

import java.io.FileOutputStream
import java.util

import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.{XSSFSheet, XSSFCell, XSSFWorkbook}
import poi.excel._
import poi.excel.api.ExcelReporter._


object SimpleReporter extends Reporter{

  val sheetName = "Avail report - All prod"

  val columnWidth = 20

  val destinationPath = "/Users/shredinger/Downloads/api_test.xlsx"

  override def run(reportInfo: ReportInfo, reportData: DataArray): ReportFile = {
    val firstColumnIndex = 0
    val firstRowIndex = 1

    val wb = new XSSFWorkbook()
    val sheet = wb.createSheet(sheetName)
    val headerRow = sheet.createRow(0)
    val dataFields = getFilteredDataFields(reportInfo).toArray
    val reportFields = getFilteredReportFields(reportInfo).toArray
    val groupFields = reportInfo.rowGroup.map(f => f.field).toSet
    val groupedCells = new util.LinkedHashMap[String, Int]()

    sheet.setDefaultColumnWidth(columnWidth)

    //header
    var columnNumber = firstColumnIndex
    for(item <- getFilteredReportFields(reportInfo)) {
      val cell = headerRow.createCell(columnNumber)
      cell.setCellValue(item.title)
      columnNumber = columnNumber + 1
    }

    //body
    var rowNumber = firstRowIndex
    for(rowItem <- prepareData(reportInfo, reportData)){
      val row = sheet.createRow(rowNumber)
      row.createCell(0).setCellValue(rowItem.head)

      var columnNumber = firstColumnIndex
      val cells = getVisibleCells(rowItem, reportInfo)
      for(cellValue <- cells){
        val cell:XSSFCell = row.createCell(columnNumber)
        val value = cellValue

        //collecting grouping data
        //work only for one group column
        if(groupFields.contains(dataFields(columnNumber).field)) collectGroupedCells(groupedCells, value)

        val dataConverter = CellDataConverter(dataFields(columnNumber).dataType)
        val columnFormatString = reportFields(columnNumber).format

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
    mergeGroupedCells(sheet, groupedCells)

    val resultFile = new FileOutputStream(destinationPath)
    wb.write(resultFile)
    resultFile.close

  }

  def mergeGroupedCells(sheet:XSSFSheet, groupedCells: util.LinkedHashMap[String, Int]) = {
    var rowNumber = 1
    val columnNumber = 0
    val it = groupedCells.values().iterator()
    while(it.hasNext){
      val cellsCount = it.next()
      sheet.addMergedRegion(new CellRangeAddress(rowNumber, rowNumber + cellsCount - 1, columnNumber, columnNumber))
      rowNumber = rowNumber + cellsCount
    }
  }

  def collectGroupedCells(cells: util.LinkedHashMap[String, Int], cellValue:String) = {
    if (!cells.containsKey(cellValue)) cells.put(cellValue, 1)
    else cells.put(cellValue, cells.get(cellValue) + 1)
  }

  def prepareData(reportInfo: ReportInfo, reportData: DataArray):DataArray = reportInfo.rowGroup match {
    case List(_) => DataSorter.sort(reportInfo, reportData)
    case Nil => reportData
  }

  def getFilteredReportFields(reportInfo: ReportInfo): List[ReportField] = {
    reportInfo.reportFields
  }

  def getFilteredDataFields(reportInfo: ReportInfo): List[DataField] = {
        val columnsToShow = reportInfo.reportFields.map(f => f.field).toSet
        reportInfo.dataFields.filter(c => columnsToShow.contains(c.field))  
  }

  def getVisibleCells(row: List[String], reportInfo: ReportInfo): List[String] = {
    val columnsToShow = reportInfo.reportFields.map(rf => rf.field).toSet
    (reportInfo.dataFields zip row)
      .filter(pair => columnsToShow.contains(pair._1.field))
      .map(pair => pair._2)
  }
}