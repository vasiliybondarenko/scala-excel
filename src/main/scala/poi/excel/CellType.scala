package poi.excel

import java.text.SimpleDateFormat
import java.util.Date

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

