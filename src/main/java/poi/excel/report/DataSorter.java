package poi.excel.report;

import poi.excel.report.api.DataField;
import poi.excel.report.api.ReportField;
import poi.excel.report.api.ReportInfo;
import poi.excel.report.util.Utils;
import poi.excel.report.util.Utils.Tuple;

import java.util.*;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/17/15
 * Time: 11:12 PM
 * Project: scala-excel
 */
public class DataSorter {
    public static List<List<String>> sort(ReportInfo reportInfo, List<List<String>> reportData){

        List<Tuple<List<Tuple<DataField, String>>, List<String>>> reportDataWithMetaData = new ArrayList<>();
        for(int i = 0; i < reportData.size(); i ++){
            List<Tuple<DataField, String>> rowWithMetaData = Utils.zip(reportInfo.getDataFields(), reportData.get(i));
            List<String> groupCells = getGroupCells(rowWithMetaData, reportInfo.getRowGroup());
            Tuple<List<Tuple<DataField, String>>, List<String>> cells = new Tuple<>(rowWithMetaData, groupCells);
            reportDataWithMetaData.add(cells);
        }
        Collections.sort(reportDataWithMetaData, new Comparator<Tuple<List<Tuple<DataField, String>>, List<String>>>() {
            @Override
            public int compare(Tuple<List<Tuple<DataField, String>>, List<String>> o1, Tuple<List<Tuple<DataField, String>>, List<String>> o2) {
                return compareLists(o1.getFiled2(), o2.getFiled2());
            }
        });

        List<List<String>> sortedReportData = new ArrayList<>();
        for(Tuple<List<Tuple<DataField, String>>, List<String>> row: reportDataWithMetaData){
            List<Tuple<DataField, String>> rowWithMeta = row.getFiled1();
            ArrayList<String> cells = new ArrayList<>();
            for(Tuple<DataField, String> tuple: rowWithMeta){
                cells.add(tuple.getFiled2());
            }
            sortedReportData.add(cells);
        }

        return sortedReportData;
    }

    private static List<String> getGroupCells(List<Utils.Tuple<DataField, String>> rowWithMeta, List<ReportField> groupFields){
        HashMap<String, String> fieldValues = new HashMap<>();
        for(Utils.Tuple<DataField, String> t: rowWithMeta){
            fieldValues.put(t.getFiled1().getField(), t.getFiled2());
        }

        ArrayList<String> cellValues = new ArrayList<>();
        for(ReportField reportField: groupFields){
            cellValues.add(fieldValues.get(reportField.getField()));
        }
        return cellValues;
    }

    private static int compareLists(List<String> l1, List<String> l2){
        for(int i = 0; i < l1.size(); i ++){
            String s1 = l1.get(i);
            String s2 = l2.get(i);
            if(!s1.equals(s2)){
                return s1.compareTo(s2);
            }
        }
        return 0;
    }
}
