package poi.excel.examples

/**
 * Created by serp on 1/13/2015.
 */
object ExcelReporterSamples {

  import poi.excel.api.ExcelReporter._

  private val acTypeCol: DataField = DataField("aircraftType", DataType.String, "AC Type")
  private val saCol: DataField = DataField("serviceArea", DataType.String, "Service Area")
  private val contractMins: DataField = DataField("minutesContracted", DataType.Number, "Contract Minutes")
  private val usedMins: DataField = DataField("minutesUsed", DataType.Number, "Used Minutes")
  private val revenue: DataField = DataField("revenue", DataType.Number, "Revenue", Option("0.00"))
  private val reportDate: DataField = DataField("reportDate", DataType.Date, "Report Date")
  private val reportMoney: DataField = DataField("reportMoney", DataType.Money, "$")

  val colInfo = List(acTypeCol, saCol, contractMins, usedMins, revenue, reportDate, reportMoney)

  val data = List(
    List("AT1", "SA1", "2000", "1000", "1000000", "2015/01/15", "2000"),
    List("AT2", "SA1", "2500", "1200", "2000000", "2015/01/14", "123000"),
    List("AT3", "SA1",  "3400", "1700", "1800000", "2015/01/12", "123000"),
    List("AT1", "SA2", "2100", "1100", "2500000", "2015/02/01", "123000"),
    List("AT2", "SA2", "2400", "1300", "2300000", "2015/01/28", "123000"),
    List("AT3", "SA2",  "3200", "1500", "2100000", "2015/01/21", "123000")
  )

  def simpleReport(): Unit = {
    //
    val repInfo = ReportInfo(colInfo,
      List(acTypeCol.field, saCol.field, contractMins.field, revenue.field, reportDate.field, reportMoney.field)
    )
    makeReporter().run(repInfo, data)
  }

  def groupedReport(): Unit = {
    // group same report by aircraft type id
    val repInfo = ReportInfo(colInfo,
      List(acTypeCol.field, saCol.field, revenue.field, reportDate.field, contractMins.field),
      List(acTypeCol.field)
    )
    makeReporter().run(repInfo, data)
  }

}
