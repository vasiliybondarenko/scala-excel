package poi.excel.api

import java.io.FileOutputStream
import java.util

import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.{XSSFCell, XSSFSheet, XSSFWorkbook}
import poi.excel._
import poi.excel.api.ExcelReporter._

import scala.collection.JavaConverters._
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
    val groupFields = reportInfo.rowGroup.map(f => f.field).toSet
    val groupedCells = new util.LinkedHashMap[Int, GroupData]()

    // Data array zipped with index
    type DataArrayZipIdx = List[(DataRow,Int)]

    // TBD: for a given row get a value by key
    def valueByField(row: DataRow, id:String) : String = null

    /**
     * Describes a node in the heirarchy group
     * @param keyField key field id
     * @param keyVal key value
     * @param children  list of children nodes
     * @param leafs list of lear row ids (only when children.isEmpty)
     */
    case class GroupNode(keyField: String, keyVal:String, children:List[GroupNode], leafs: List[Int])

    // data fields zipped with index
    val dataFieldsZipIdx:DataArrayZipIdx = reportData.zipWithIndex



    val groupsNodes:List[GroupNode] = groupFields.toList.map { fld =>
      // TODO must be implemeneted
      null
    }




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
        if(groupFields.contains(dataFields(columnNumber).field))
          collectGroupedCells(groupedCells, value, columnNumber, rowNumber)

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

  //Incorrect
  def mergeGroupedCells(sheet:XSSFSheet, groupedCells: util.LinkedHashMap[Int, GroupData]) = {
    val it = groupedCells.values().iterator()
    var prevColumnGroupData:GroupData = null
    while (it.hasNext){
      val groupData = it.next()
      var rowNumber = 1
      val columnNumber = groupData.columnIndex
      var prevCount = groupData.counts.get(0)
      for(count <- groupData.counts.asScala){
        val currentCount = getGroupCount(prevColumnGroupData, rowNumber, columnNumber, count)
        println((rowNumber, columnNumber, currentCount, count))
        sheet.addMergedRegion(new CellRangeAddress(rowNumber,  rowNumber + currentCount - 1, columnNumber, columnNumber))
        rowNumber = rowNumber + currentCount
        prevCount = count
      }
      prevColumnGroupData = groupData
    }
  }

  def getGroupCount(prevGroupData: GroupData, row:Int, col:Int, count:Int):Int = {
    def findPrevRow(counts:List[Int], row:Int):(Int, Int) = {
      var prevColRow = 0
      for(c <- counts){
        if(row >= prevColRow)  return (prevColRow, c) else
          prevColRow = prevColRow + c - 1
      }
      prevColRow -> 0
    }

    if(prevGroupData == null) count else {
      val counts = prevGroupData.counts.asScala
      val (_, prevColRowCount) = findPrevRow(prevGroupData.counts.asScala.toList, row)
      prevColRowCount min count
    }
  }

  def collectGroupedCells(cells: util.LinkedHashMap[Int, GroupData], cellValue:String, column:Int, row:Int) = {
    if (!cells.containsKey(column)) cells.put(column, GroupData(column))
    cells.get(column).collect(cellValue, row)
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

  case class GroupData(columnIndex: Int){
    private val cellValues = new util.ArrayList[String]()
    val counts = new util.ArrayList[Int]()

    def collect(value:String, rowIndex:Int) = {
      if(cellValues.isEmpty) {
        cellValues.add(value)
        counts.add(1)
      } else {
        if(cellValues.get(cellValues.size() - 1) != value){
          cellValues.add(value)
          counts.add(1)
        } else {
          val lastIndex = counts.size() - 1
          cellValues.add(value)
          counts.set(lastIndex, counts.get(lastIndex) + 1)
        }
      }
    }
  }
}