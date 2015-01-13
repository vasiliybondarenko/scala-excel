package poi.excel.examples

import poi.excel.api._

/**
 * Created by serp on 1/13/2015.
 */
object ExcelReporterSamples {

  import poi.excel.api.ExcelReporter._

  private val acTypeCol: ColInfo = ColInfo("aircraftType", DataType.String, "AC Type")
  private val saCol: ColInfo = ColInfo("serviceArea", DataType.String, "Service Area")
  private val contractMins: ColInfo = ColInfo("minutesContracted", DataType.Number, "Contract Minutes")
  private val usedMins: ColInfo = ColInfo("minutesUsed", DataType.Number, "Used Minutes")
  private val revenue: ColInfo = ColInfo("revenue", DataType.Number, "Revenue")
  private val reportDate: ColInfo = ColInfo("reportDate", DataType.Date, "Report Date")

  val data = List(
    List("AT1", "SA1", "2000", "1000", "1000000", "2015/01/15"),
    List("AT2", "SA1", "2500", "1200", "2000000", "2015/01/14"),
    List("AT3", "SA1",  "3400", "1700", "1800000", "2015/01/12"),
    List("AT1", "SA2", "2100", "1100", "2500000", "2015/02/01"),
    List("AT2", "SA2", "2400", "1300", "2300000", "2015/01/28"),
    List("AT3", "SA2",  "3200", "1500", "2100000", "2015/01/21")
  )

  def simpleReport(): Unit = {
    val repInfo = ReportInfo(List(acTypeCol, saCol, contractMins, revenue, reportDate))
    makeReporter().run(repInfo, data)
  }

  def groupedReport(): Unit = {
    // group same report by aircraft type id
    val repInfo = ReportInfo(List(acTypeCol, saCol, revenue, reportDate, contractMins), List(acTypeCol.field))
    makeReporter().run(repInfo, data)
  }

}
