package poi.excel.api

import java.io.FileOutputStream

import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.{XSSFCell, XSSFWorkbook}
import poi.excel._
import poi.excel.api.ExcelReporter._
//import scala.collection.JavaConverters._

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

    sheet.setDefaultColumnWidth(columnWidth)

    //header
    var columnNumber = firstColumnIndex
    val filteredReportFields = getFilteredReportFields(reportInfo)
    for(item <- filteredReportFields) {
      val cell = headerRow.createCell(columnNumber)
      cell.setCellValue(item.title)
      columnNumber = columnNumber + 1
    }

    //body
    var rowNumber = firstRowIndex
    val preparedData: DataArray = prepareData(reportInfo, reportData)
    for(rowItem <- preparedData){
      val row = sheet.createRow(rowNumber)
      row.createCell(0).setCellValue(rowItem.head)

      var columnNumber = firstColumnIndex
      val cells = getVisibleCells(rowItem, reportInfo)
      for(cellValue <- cells){
        val cell:XSSFCell = row.createCell(columnNumber)
        val value = cellValue

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
    reportInfo.rowGroup match {
      case Nil =>
      case _ => {
        val groupNodes = GroupHelper.createGroupNodes(reportInfo, preparedData)
        val mergeRegions = GroupHelper.getMergeRegions(groupNodes)
        val fieldsWithIndexes = filteredReportFields.map(f => f.field).zipWithIndex.toMap
        mergeRegions.foreach(mergeRegion => {
          val columnIndex = fieldsWithIndexes.get(mergeRegion.field).get
          sheet.addMergedRegion(new CellRangeAddress(firstRowIndex + mergeRegion.firstRow, firstRowIndex + mergeRegion.lastRow, columnIndex, columnIndex))
        })
      }
    }

    val resultFile = new FileOutputStream(destinationPath)
    wb.write(resultFile)
    resultFile.close

  }

  def prepareData(reportInfo: ReportInfo, reportData: DataArray):DataArray = reportInfo.rowGroup match {
    case Nil => reportData
    case _ => DataSorter.sort(reportInfo, reportData)
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