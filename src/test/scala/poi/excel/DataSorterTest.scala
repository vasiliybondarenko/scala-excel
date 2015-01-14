package poi.excel

import org.scalatest.{FlatSpec, Matchers}
import poi.excel.api.ExcelReporter.{ReportInfo, DataType, ColInfo}

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/14/15
 * Time: 1:18 PM 
 * Project: scala-excel
 */
class DataSorterTest extends FlatSpec with Matchers {

  private val acTypeCol: ColInfo = ColInfo("aircraftType", DataType.String, "AC Type")
  private val saCol: ColInfo = ColInfo("serviceArea", DataType.String, "Service Area")
  private val usedMins: ColInfo = ColInfo("minutesUsed", DataType.Number, "Used Minutes")

  val data = List(
    List("AT1", "SA1"),
    List("AT2", "SA1"),
    List("AT3", "SA1"),
    List("AT1", "SA2"),
    List("AT2", "SA2"),
    List("AT3", "SA2")
  )

  val data1 = List(
    List("AT1", "SA1", "6"),
    List("AT1", "SA2", "5"),
    List("AT1", "SA3", "4"),
    List("AT2", "SA1", "3"),
    List("AT2", "SA2", "2"),
    List("AT2", "SA3", "1")
  )

  it should "sort data by group column" in {
    DataSorter.sort(ReportInfo(List(acTypeCol, saCol), List(saCol.field, saCol.field), List(acTypeCol.field)), data) should be(
      List(
        List("AT1", "SA1"),
        List("AT1", "SA2"),
        List("AT2", "SA1"),
        List("AT2", "SA2"),
        List("AT3", "SA1"),
        List("AT3", "SA2")
      )
    )
  }

  it should "sort data by specified non-first group column" in {
    DataSorter.sort(
      ReportInfo(
        List(acTypeCol, saCol, usedMins),
        List(saCol.field, usedMins.field),
        List(saCol.field)),
      data1) should be(
      List(
        List("AT1", "SA1", "6"),
        List("AT2", "SA1", "3"),
        List("AT1", "SA2", "5"),
        List("AT2", "SA2", "2"),
        List("AT1", "SA3", "4"),
        List("AT2", "SA3", "1")
      )
    )
  }

  //not implemented
  ignore  should "sort data by > 1 group columns" in {
    DataSorter.sort(
      ReportInfo(
        List(acTypeCol, saCol, usedMins),
        List(saCol.field, usedMins.field),
        List(saCol.field, usedMins.field)),
      data1) should be(
      List(
        List("AT2", "SA1", "3"),
        List("AT1", "SA1", "6"),
        List("AT2", "SA2", "2"),
        List("AT1", "SA2", "5"),
        List("AT2", "SA3", "1"),
        List("AT1", "SA3", "4")
      )
    )
  }

}
