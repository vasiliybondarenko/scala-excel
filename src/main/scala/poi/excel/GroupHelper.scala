package poi.excel

import poi.excel.api.ExcelReporter._

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/15/15
 * Time: 4:39 PM 
 * Project: scala-excel
 */
object GroupHelper {
  // Data array zipped with index
  type DataArrayZipIdx = List[(DataRow,Int)]

  /**
   * Describes a node in the heirarchy group
   * @param keyField key field id
   * @param keyVal key value
   * @param children  list of children nodes
   * @param leafs list of lear row ids (only when children.isEmpty)
   */
  case class GroupNode(keyField: String, keyVal:String, children:List[GroupNode], leafs: List[Int])

  // TBD: for a given row get a value by key
  def valueByField(row: DataRow, id:String) : String = null

  def createGroupNodes(reportInfo: ReportInfo, reportData: DataArray): List[GroupNode] = {
    // data fields zipped with index
    val dataFieldsZipIdx:DataArrayZipIdx = reportData.zipWithIndex

    val groupFields = reportInfo.rowGroup.map(f => f.field)


    val groupsNodes:List[GroupNode] = groupFields.map { fld =>
      // TODO must be implemeneted
      null
    }

    groupsNodes
  }



}
