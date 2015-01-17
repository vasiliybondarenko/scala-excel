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

  case class MergeRegion(field:String, firstRow:Int, lastRow:Int)

  def createGroupNodes(reportInfo: ReportInfo, reportData: DataArray): List[GroupNode] = {
    // data fields zipped with index
    val dataFieldsZipIdx:DataArrayZipIdx = reportData.zipWithIndex

    // TBD: for a given row get a value by key
    //TODO: optimize
    lazy val dataFieldIndexById: Map[String, Int] = (reportInfo.dataFields.map(x => x.field) zipWithIndex).toMap
    def valueByField(row: DataRow, id:String) : String = {
      row(dataFieldIndexById.get(id).get)
    }

    val groupFields = reportInfo.rowGroup.map(f => f.field)

    val groupsNodes:List[GroupNode] = groupFields.map { fld =>
      // TODO must be implemeneted
      null
    }

    def createGroupNodes(dataArray:DataArrayZipIdx, groupFields:List[String]):List[GroupNode] = {

      dataArray.groupBy(record => valueByField(record._1, groupFields.head))
        .map(x => GroupNode(groupFields.head, x._1,
          groupFields.tail match {
            case Nil => List()
            case _ => createGroupNodes(x._2, groupFields.tail)
          },
          groupFields.tail match {
            case Nil => x._2.map(dataRowWithIndex => dataRowWithIndex._2)
            case _ => Nil
          })
        )
        .toList.sortWith((l, r) => l.keyVal < r.keyVal)
    }

    createGroupNodes(dataFieldsZipIdx, groupFields)
  }

  def getMergeRegions(groupNodes: List[GroupNode]):List[MergeRegion] = {
    groupNodes.map(groupNode =>
      groupNode.leafs match {
        case Nil => {
          val leaves = getAllLeaves(groupNode.children)
          MergeRegion(groupNode.keyField, leaves.min, leaves.max) :: getMergeRegions(groupNode.children)
        }
        case _ => MergeRegion(groupNode.keyField, groupNode.leafs.min, groupNode.leafs.max) :: Nil
      }).flatten
  }

  private def getAllLeaves(children: List[GroupNode]): List[Int] = {
    children.map(node => node.leafs match {
      case Nil => getAllLeaves(node.children)
      case _ => node.leafs
    }).flatten.sorted
  }
}
