package poi.excel

import org.scalatest.{FlatSpec, Matchers}
import poi.excel.api.ExcelReporter._

/**
* Created with IntelliJ IDEA.
* Author: shredinger
* Date: 1/14/15
* Time: 1:18 PM
* Project: scala-excel
*/
class DataSorterTest extends FlatSpec with Matchers {

  private val acType = DataField("aircraftType", DataType.String)
  private val sa = DataField("serviceArea", DataType.String)
  private val usedMins = DataField("minutesUsed", DataType.Number)


  private val acTypeRF = ReportField("aircraftType", "AC Type")
  private val saRF = ReportField("serviceArea", "Service Area")
  private val usedMinsRF = ReportField("minutesUsed", "Used Minutes")

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

  val data2 = List(
    List("AT3", "SA3", "6"),
    List("AT2", "SA2", "5"),
    List("AT1", "SA3", "4"),
    List("AT3", "SA1", "3"),
    List("AT2", "SA1", "2"),
    List("AT1", "SA2", "1")
  )

  it should "sort data by group column" in {
    DataSorter.sort(ReportInfo(List(acType, sa), List(acTypeRF, saRF), List(acTypeRF)), data) should be(
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
        List(acType, sa, usedMins),
        List(saRF, usedMinsRF),
        List(saRF)),
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
  it  should "sort data by > 1 group columns" in {
    DataSorter.sort(
      ReportInfo(
        List(acType, sa, usedMins),
        List(acTypeRF, saRF),
        List(acTypeRF, saRF)),
      data2) should be(
      List(
        List("AT1", "SA2", "1"),
        List("AT1", "SA3", "4"),
        List("AT2", "SA1", "2"),
        List("AT2", "SA2", "5"),
        List("AT3", "SA1", "3"),
        List("AT3", "SA3", "6")
      )
    )
  }

}
