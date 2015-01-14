package poi.excel.api

/**
 * Created by serp on 1/13/2015.
 */
object ExcelReporter {


  /**
   * Internal data type
   */
  object DataType extends Enumeration {
    val
      Number,
      Money,
      String,
      Date // YYYY/MM/DD
    = Value
  }
  
  /**
   * 
   * @param field   data field name (map key)
   */
  case class DataField(field: String, dataType: DataType.Value)

  /**
   *
   * @param field
   * @param title
   * @param format  format string to visualize
   */
  case class ReportField(field: String, title:String, format: Option[String] = None)

  // meta column list
  type DataFieldList = List[DataField]

  // list of strings
  type GroupFieldList = List[ReportField]

  type ReportFieldList = List[ReportField]

  // 2 dims report data
  type DataRow = List[String]
  type DataArray = List[DataRow]

  // report info
  case class ReportInfo( dataFields: DataFieldList, reportFields: ReportFieldList, rowGroup:GroupFieldList = Nil )

  type ReportFile = Any // TDB provide the actual data type

  trait Reporter {
    def run( reportInfo: ReportInfo, reportData: DataArray) : ReportFile
  }

  def makeReporter() : Reporter = {
    return new Reporter {
      override def run(reportInfo: ReportInfo, reportData: DataArray): ReportFile = {
        // TDB write the actual code here
        SimpleReporter.run(reportInfo, reportData)
      }
    }
  }


}
