package poi.excel

import java.text.SimpleDateFormat
import java.util.Date

import org.apache.poi.xssf.usermodel.XSSFCell
import poi.excel.api.ExcelReporter

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/12/15
 * Time: 11:03 PM 
 * Project: scala-excel
 */
object CellDataTypes{
  val CURRENCY = "^\\d*\\.\\d*$"
  val DATE = "\\d{4}"
}


abstract class CellData[+T]() {
  def formatString:String
  def getValue:T
}

case class Currency(value:String) extends CellData[Double]{
  override def formatString: String = "$#,##0.00"
  override def getValue: Double = value.toDouble
}

case class  DateCellType(value:String) extends CellData[Date]{
  val formatter = new SimpleDateFormat(formatString)
  override def getValue: Date = formatter.parse(value)
  override def formatString: String = "dd/mm/yyyy"
}

case class Text(value:String) extends CellData[String]{
  override def formatString: String = ""
  override def getValue: String = value
}

//for new API

abstract class CellDataConverter{
  def getDataFormat:String
  def assignValue(value:String, cell: XSSFCell)
}

case class NumberConverter(dataType:ExcelReporter.DataType.Value) extends CellDataConverter{
  override def getDataFormat: String = "0.0000"
  override def assignValue(value: String, cell: XSSFCell): Unit = cell.setCellValue(value.toDouble)
}

case class DateConverter(dataType:ExcelReporter.DataType.Value) extends CellDataConverter{
  val formatter = new SimpleDateFormat("yyyy/mm/dd")
  override def getDataFormat: String = "dd/mm/yy"
  override def assignValue(value: String, cell: XSSFCell): Unit = cell.setCellValue(formatter.parse(value))
}

case class TextConverter(dataType:ExcelReporter.DataType.Value) extends CellDataConverter{
  override def getDataFormat: String = ""
  override def assignValue(value: String, cell: XSSFCell): Unit = cell.setCellValue(value)
}

object CellDataConverter{
  def apply(dataType:ExcelReporter.DataType.Value):CellDataConverter = dataType match {
    case ExcelReporter.DataType.Number => NumberConverter(dataType)
    case ExcelReporter.DataType.Date => DateConverter(dataType)
    case ExcelReporter.DataType.String => TextConverter(dataType)
  }
}
