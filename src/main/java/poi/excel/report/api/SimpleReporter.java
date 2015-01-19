package poi.excel.report.api;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import poi.excel.report.DataSorter;
import poi.excel.report.GroupHelper;
import poi.excel.report.GroupNode;
import poi.excel.report.MergeRegion;
import poi.excel.report.converters.CellDataConverter;
import poi.excel.report.converters.CellDataConverterFactory;
import poi.excel.report.util.Utils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/16/15
 * Time: 4:36 PM
 * Project: scala-excel
 */
public class SimpleReporter implements Reporter{

    private final String sheetName = "Avail report - All prod";

    private final int columnWidth = 20;

    private final String destinationPath = "/Users/shredinger/Downloads/java_port_test.xlsx";

    @Override
    public void run(ReportInfo reportInfo, List<List<String>> reportData) {
        final int firstColumnIndex = 0;
        final int firstRowIndex = 1;

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName);
        XSSFRow headerRow = sheet.createRow(0);
        List<DataField> dataFields = getFilteredDataFields(reportInfo);
        List<ReportField> reportFields = getFilteredReportFields(reportInfo);

        //header
        int columnNumber = firstColumnIndex;
        List<ReportField> filteredReportFields = getFilteredReportFields(reportInfo);
        for(ReportField item: filteredReportFields) {
            XSSFCell cell = headerRow.createCell(columnNumber);
            cell.setCellValue(item.getTitle());
            columnNumber = columnNumber + 1;
        }

        //body
        List<List<String>> preparedData = prepareData(reportInfo, reportData);
        int rowNumber = firstRowIndex;
        for(int r = 0; r < preparedData.size(); r ++){
            List<String> cells = getVisibleCells(preparedData.get(r), reportInfo);
            XSSFRow row = sheet.createRow(r + firstRowIndex);
            row.createCell(0).setCellValue(cells.get(0));


            for (int col = 0; col < cells.size(); col ++){
                XSSFCell cell = row.createCell(firstColumnIndex + col);

                DataType dataType = dataFields.get(col).getDataType();
                CellDataConverter cellDataConverter = CellDataConverterFactory.create(dataType);
                String columnFormatString = reportFields.get(col).getFormat();
                XSSFCellStyle style = wb.createCellStyle();
                XSSFDataFormat format = wb.createDataFormat();
                String formatString = columnFormatString != null ? columnFormatString : cellDataConverter.getDataFormat();
                style.setDataFormat(format.getFormat(formatString));
                cell.setCellStyle(style);
                cellDataConverter.assignValue(cells.get(col), cell);
            }
        }

        //grouping
        if(!reportInfo.getRowGroup().isEmpty()){
            List<GroupNode> groupNodes = GroupHelper.createGroupNodes(reportInfo, preparedData);
            List<MergeRegion> mergeRegions = GroupHelper.getMergeRegions(groupNodes);

            Map<String, Integer> fieldsWithIndexes = getFieldsWithIndexes(filteredReportFields);
            for (MergeRegion mergeRegion: mergeRegions){
                Integer colIndex = fieldsWithIndexes.get(mergeRegion.getField());
                int firstRow = mergeRegion.getFirstRow() + firstRowIndex;
                int lastRow = mergeRegion.getLastRow() + firstRowIndex;
                sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, colIndex, colIndex));
            }
        }

        sheet.setDefaultColumnWidth(columnWidth);

        FileOutputStream resultFile = null;
        try {
            resultFile = new FileOutputStream(destinationPath);
            wb.write(resultFile);
            resultFile.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    private Map<String, Integer> getFieldsWithIndexes(List<ReportField> reportFields){
        HashMap<String, Integer> fieldsWithIndexes = new HashMap<>();
        for (int i = 0; i < reportFields.size(); i++) {
            fieldsWithIndexes.put(reportFields.get(i).getField(), i);
        }
        return fieldsWithIndexes;
    }

    private List<String> getVisibleCells(List<String> rowData, ReportInfo reportInfo){
        Set<String> reportFields = new HashSet<String>();
        for(ReportField rf: Utils.toSet(reportInfo.getReportFields())){
            reportFields.add(rf.getField());
        }
        ArrayList<String> cells = new ArrayList<>();
        for(int i = 0; i < rowData.size(); i ++){
            if(reportFields.contains(reportInfo.getDataFields().get(i).getField())){
                cells.add(rowData.get(i));
            }
        }
        return cells;
    }

    private List<List<String>> prepareData(ReportInfo reportInfo, List<List<String>> reportData){
        if(reportInfo.getRowGroup().isEmpty())
            return reportData;
        else return DataSorter.sort(reportInfo, reportData);
    }

    private List<DataField> getFilteredDataFields(ReportInfo reportInfo){
        Set<String> reportFields = new HashSet<String>();
        for(ReportField rf: Utils.toSet(reportInfo.getReportFields())){
            reportFields.add(rf.getField());
        }
        ArrayList<DataField> dataFields = new ArrayList<>();
        for(DataField dataField: reportInfo.getDataFields()){
            if(reportFields.contains(dataField.getField())){
                dataFields.add(dataField);
            }
        }
        return dataFields;
    }

    private List<ReportField> getFilteredReportFields(ReportInfo reportInfo){
        return reportInfo.getReportFields();
    }
}
