package poi.excel.report;

import com.google.common.collect.ArrayListMultimap;
import poi.excel.report.api.DataField;
import poi.excel.report.api.ReportField;
import poi.excel.report.api.ReportInfo;
import poi.excel.report.util.Mapper;
import poi.excel.report.util.Utils;
import poi.excel.report.util.Utils.Tuple;

import java.util.*;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/18/15
 * Time: 9:14 PM
 * Project: scala-excel
 */
public class GroupHelper {

    public static List<GroupNode> createGroupNodes(ReportInfo reportInfo, List<List<String>> reportData){


        Map<String, Integer> dataFieldIndexById = new HashMap<>();
        List<DataField> dataFields = reportInfo.getDataFields();
        for(int i = 0; i < dataFields.size(); i ++){
            dataFieldIndexById.put(dataFields.get(i).getField(), i);
        }

        List<String> groupFields = Utils.map(reportInfo.getRowGroup(), new Mapper<ReportField, String>() {
            @Override
            public String apply(ReportField value) {
                return value.getField();
            }
        });

        List<Utils.Tuple<List<String>, Integer>> rowsWithIndex = zipWithIndex(reportData);

        return createGroupNodes(rowsWithIndex, dataFieldIndexById, groupFields);
    }

    public static List<MergeRegion> getMergeRegions(List<GroupNode> groupNodes){
        ArrayList<MergeRegion> mergeRegions = new ArrayList<>();
        for(GroupNode groupNode: groupNodes){
            List<Integer> leaves = groupNode.getLeaves();
            if(leaves.isEmpty()){
                List<Integer> allLeaves = getAllLeaves(groupNode.getChildren());
                int minRow = Collections.min(allLeaves);
                int maxRow = Collections.max(allLeaves);
                mergeRegions.add(new MergeRegion(groupNode.getKey(), minRow, maxRow));
                mergeRegions.addAll(getMergeRegions(groupNode.getChildren()));
            } else {
                int firstRow = Collections.min(leaves);
                int lastRow = Collections.max(leaves);
                mergeRegions.add(new MergeRegion(groupNode.getKey(), firstRow, lastRow));
            }
        }

        return mergeRegions;
    }

    private static List<Integer> getAllLeaves(List<GroupNode> children){
        ArrayList<Integer> leaves = new ArrayList<>();
        for(GroupNode groupNode: children){
            if(groupNode.getLeaves().isEmpty()){
                leaves.addAll(getAllLeaves(groupNode.getChildren()));
            }else{
                leaves.addAll(groupNode.getLeaves());
            }
        }
        return leaves;
    }

    private static List<GroupNode> createGroupNodes(Collection<Tuple<List<String>, Integer>> dataArray,
                                                    Map<String, Integer> dataFieldsByIndex,
                                                    List<String> groupFields){
        ArrayList<GroupNode> groupNodes = new ArrayList<>();
        String groupField = groupFields.get(0);
        Map<String, Collection<Tuple<List<String>, Integer>>> groupedDataArray = group(dataArray, dataFieldsByIndex, groupField);
        for(String value: groupedDataArray.keySet()){
            Collection<Tuple<List<String>, Integer>> rowsForKey = groupedDataArray.get(value);


            List<Integer> leaves = groupFields.size() == 1 ? getRowIndexes(rowsForKey) : new ArrayList<Integer>();
            List<GroupNode> children = groupFields.size() == 1 ?
                    new ArrayList<GroupNode>() :
                    createGroupNodes(rowsForKey, dataFieldsByIndex, groupFields.subList(1, groupFields.size()));
            groupNodes.add(new GroupNode(groupField, value, children, leaves));
        }

        sortGroupNodes(groupNodes);

        return groupNodes;
    }

    private static List<GroupNode> sortGroupNodes(List<GroupNode> groupNodes){
        Collections.sort(groupNodes, new Comparator<GroupNode>() {
            @Override
            public int compare(GroupNode o1, GroupNode o2) {
                return o1.getValue().compareTo(o2.getValue());
            }
        });
        return groupNodes;
    }

    private static List<Integer> getRowIndexes(Collection<Tuple<List<String>, Integer>> rows){
        ArrayList<Integer> rowIndexes = new ArrayList<>();
        for (Tuple<List<String>, Integer> rowWithIndexes: rows){
            rowIndexes.add(rowWithIndexes.getFiled2());
        }
        return rowIndexes;
    }

    private static Map<String, Collection<Tuple<List<String>, Integer>>> group(Collection<Tuple<List<String>, Integer>> dataArray,
                                                                                      Map<String, Integer> dataFieldIndexById,
                                                                                      String groupField){
        ArrayListMultimap<String, Tuple<List<String>, Integer>> groupedRows = ArrayListMultimap.create();
        for(Tuple<List<String>, Integer> rowWithIndex: dataArray){
            Integer indexInRow = dataFieldIndexById.get(groupField);
            String cellValue = rowWithIndex.getFiled1().get(indexInRow);
            groupedRows.put(cellValue, rowWithIndex);
        }
        return groupedRows.asMap();

    }

    private static List<Utils.Tuple<List<String>, Integer>> zipWithIndex(List<List<String>> dataArray){
        ArrayList<Utils.Tuple<List<String>, Integer>> rowsWithIndex = new ArrayList<>();
        for(int row = 0; row < dataArray.size(); row ++){
            rowsWithIndex.add(new Utils.Tuple<List<String>, Integer>(dataArray.get(row), row));
        }
        return rowsWithIndex;
    }


}
