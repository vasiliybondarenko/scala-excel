package poi.excel

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/12/15
 * Time: 11:03 PM 
 * Project: scala-excel
 */
abstract class CellData() {
  def formatString:String
}

case class Number(value:String) extends CellData{
  override def formatString: String = "#,##0.00"
}
