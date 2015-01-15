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
  def sort(reportInfo: ReportInfo, reportData: DataArray):DataArray
}

class DataSorter extends Sorter{
  override def sort(reportInfo: ReportInfo, reportData: DataArray): DataArray = {

    def getGroupCells(rowWithMeta: List[(DataField, String)], groupFields: List[ReportField]): List[String] = {
      for(groupColumn <- groupFields)
        yield rowWithMeta.find(cellWithMeta => cellWithMeta._1.field == groupColumn.field).get._2
    }

    def compare(l:List[String], r:List[String]):Boolean = (l, r) match {
      case (h1::t1, h2::t2) => if(h1 == h2) compare(t1, t2) else h1 < h2
      case (List(h1), List(h2)) => h1 < h2
      case (Nil, Nil) => false
    }

    reportData.map(row => (reportInfo.dataFields zip row))
      .map(rowWithMeta => (rowWithMeta, getGroupCells(rowWithMeta, reportInfo.rowGroup)))
      .sortWith((l, r) => compare(l._2, r._2))
      .map(row => row._1).map(row => row.map(record => record._2))

  }
}

object DataSorter extends DataSorter{
  def apply = new DataSorter
}
