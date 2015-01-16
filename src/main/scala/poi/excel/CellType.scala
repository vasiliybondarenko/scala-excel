package poi.excel

import java.text.SimpleDateFormat

import org.apache.poi.xssf.usermodel.XSSFCell
import poi.excel.api.ExcelReporter

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/12/15
 * Time: 11:03 PM 
 * Project: scala-excel
 */

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

case class MoneyConverter(dataType:ExcelReporter.DataType.Value) extends CellDataConverter{
  override def getDataFormat: String = "$#,##0.00"
  override def assignValue(value: String, cell: XSSFCell): Unit = cell.setCellValue(value.toDouble)
}

object CellDataConverter{
  def apply(dataType:ExcelReporter.DataType.Value):CellDataConverter = dataType match {
    case ExcelReporter.DataType.Number => NumberConverter(dataType)
    case ExcelReporter.DataType.Date => DateConverter(dataType)
    case ExcelReporter.DataType.Money => MoneyConverter(dataType)
    case ExcelReporter.DataType.String => TextConverter(dataType)
  }
}
