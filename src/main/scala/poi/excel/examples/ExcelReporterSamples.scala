package poi.excel.examples

/**
 * Created by serp on 1/13/2015.
 */
object ExcelReporterSamples {

  import poi.excel.api.ExcelReporter._

  private val acTypeCol: DataField = DataField("aircraftType", DataType.String)
  private val saCol: DataField = DataField("serviceArea", DataType.String)
  private val contractMins: DataField = DataField("minutesContracted", DataType.Number)
  private val usedMins: DataField = DataField("minutesUsed", DataType.Number)
  private val revenue: DataField = DataField("revenue", DataType.Number)
  private val reportDate: DataField = DataField("reportDate", DataType.Date)
  private val reportMoney: DataField = DataField("reportMoney", DataType.Money)
  
  private val acTypeRepField = ReportField("aircraftType", "AC Type")
  private val saRF = ReportField("serviceArea", "Service Area")
  private val contractMinsRF = ReportField("minutesContracted", "Contract Minutes")
  private val usedMinsRF = ReportField("minutesUsed", "Used Minutes")
  private val revenueRF = ReportField("revenue", "Revenue", Option("0.00"))
  private val reportDateRF = ReportField("reportDate", "Report Date")
  private val reportMoneyRF = ReportField("reportMoney", "$")

  val dataFields = List(acTypeCol, saCol, contractMins, usedMins, revenue, reportDate, reportMoney)

  val data = List(
    List("AT1", "SA2", "2000", "1000", "1000000", "2015/01/15", "2000"),
    List("AT2", "SA2", "2500", "1200", "2000000", "2015/01/14", "123000"),
    List("AT3", "SA2",  "3400", "1700", "1800000", "2015/01/12", "123000"),
    List("AT1", "SA2", "2100", "1100", "2500000", "2015/02/01", "123000"),
    List("AT2", "SA1", "2400", "1300", "2300000", "2015/01/28", "123000"),
    List("AT3", "SA2",  "3200", "1500", "2100000", "2015/01/21", "123000")
  )

  def simpleReport(): Unit = {
    //
    val repInfo = ReportInfo(dataFields,
      List(acTypeRepField, saRF, contractMinsRF, revenueRF, reportDateRF, reportMoneyRF)
    )
    makeReporter().run(repInfo, data)
  }

  def groupedReport(): Unit = {
    // group same report by aircraft type id
    val repInfo = ReportInfo(dataFields,
      List(acTypeRepField, saRF, contractMinsRF, revenueRF, reportDateRF),
      List(acTypeRepField, saRF)
    )
    makeReporter().run(repInfo, data)
  }

}
