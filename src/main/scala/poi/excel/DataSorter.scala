package poi.excel

import poi.excel.api.ExcelReporter._

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/14/15
 * Time: 1:09 PM 
 * Project: scala-excel
 */

trait Sorter{
  def sort(reportInfo: ReportInfo, reportData: ReportData):ReportData
}

class DataSorter extends Sorter{
  override def sort(reportInfo: ReportInfo, reportData: ReportData): ReportData = {
    def compare(l1:List[String], l2:List[String]) = l1.head < l2.head

    //
    val groupColumn = reportInfo.rowGroup.head
    reportData.map(row => (reportInfo.colInfo zip row))
      .map(rowWithMeta => (rowWithMeta, rowWithMeta.find(cellWithMeta => cellWithMeta._1.field == groupColumn).get._2))
      .sortWith((l1, l2) => l1._2 < l2._2)
      .map(row => row._1).map(row => row.map(record => record._2))
    //

    //reportData.sortWith(compare)
  }
}

object DataSorter extends DataSorter{
  def apply = new DataSorter
}
