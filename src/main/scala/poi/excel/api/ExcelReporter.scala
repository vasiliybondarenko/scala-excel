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
      String,
      Date // YYYY/MM/DD
    = Value
  }
  
  /**
   * 
   * @param field   data field name (map key)
   * @param title   title 
   * @param dataType data type               
   * @param format  format string to visualize
   */
  case class ColInfo(field: String, dataType: DataType.Value, title:String, format: Option[String] = None )

  // meta column list
  type ColInfoList = List[ColInfo]

  // list of strings
  type RowGroupList = List[String]

  type ShowColList = List[String]

  // 2 dims report data
  type ReportData = List[List[String]]

  // report info
  case class ReportInfo( colInfo: ColInfoList, colShowList: ShowColList, rowGroup:RowGroupList = Nil )

  type ReportFile = Any // TDB provide the actual data type

  trait Reporter {
    def run( reportInfo: ReportInfo, reportData: ReportData) : ReportFile
  }

  def makeReporter() : Reporter = {
    return new Reporter {
      override def run(reportInfo: ReportInfo, reportData: ReportData): ReportFile = {
        // TDB write the actual code here
      }
    }
  }


}
