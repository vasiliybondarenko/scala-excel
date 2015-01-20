package poi.excel.report.api;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import poi.excel.report.*;
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
public class SimpleReporter implements Reporter, ReportStyle{


    private final int columnWidth = 20;

    private final String destinationPath = "/Users/shredinger/Downloads/java_port_test.xlsx";

    @Override
    public void run(ReportInfo reportInfo, List<List<String>> reportData) {
        final int firstColumnIndex = 0;
        final int firstRowIndex = 4;

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(reportInfo.getSpreadshitName());
        XSSFRow headerRow = sheet.createRow(firstRowIndex - 1);
        List<DataField> dataFields = getFilteredDataFields(reportInfo);
        List<ReportField> reportFields = getFilteredReportFields(reportInfo);

        //title/logo
        int titleStartColumnIndex = 2;
        int titleStartRowIndex = 1;
        short titleRowHeight = (short)500;
        XSSFRow titleRow = sheet.createRow(titleStartRowIndex);
        titleRow.setHeight(titleRowHeight);
        XSSFCell titleCell = titleRow.createCell(titleStartColumnIndex);
        titleCell.setCellValue(reportInfo.getTitleName());
        titleCell.setCellStyle(getTitleRowStyle(wb));


        //header
        int columnNumber = firstColumnIndex;
        List<ReportField> filteredReportFields = getFilteredReportFields(reportInfo);
        XSSFCellStyle headerRowStyle = getHeaderRowStyle(wb);
        for(ReportField item: filteredReportFields) {
            XSSFCell cell = headerRow.createCell(columnNumber);
            cell.setCellValue(item.getTitle());
            columnNumber = columnNumber + 1;
            cell.setCellStyle(headerRowStyle);
        }

        //body
        List<List<String>> preparedData = prepareData(reportInfo, reportData);
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

    private XSSFCellStyle getStyle(XSSFWorkbook wb, CellFontInfo cellStyleInfo){
        XSSFCellStyle style = wb.createCellStyle();
        XSSFFont font = wb.createFont();

        if(cellStyleInfo.fontSize != null){
            font.setFontHeightInPoints(cellStyleInfo.fontSize);
        }

        if(cellStyleInfo.fontColor != null){
            font.setColor(cellStyleInfo.fontColor);
        }

        if(cellStyleInfo.boldWeight != null){
            font.setBoldweight(cellStyleInfo.boldWeight);
        }

        if(cellStyleInfo.backgroundColor != null){
            style.setFillBackgroundColor(cellStyleInfo.backgroundColor);
            style.setFillPattern(CellStyle.BIG_SPOTS);
        }

        style.setFont(font);
        return style;
    }

    private XSSFCellStyle getTitleRowStyle(XSSFWorkbook wb){
        return getStyle(wb, new CellFontInfo(TITLE_FONT_COLOR, TITLE_FONT_SIZE, null, BOLD_WEIGHT));
    }

    private XSSFCellStyle getHeaderRowStyle(XSSFWorkbook wb){
        return getStyle(wb, new CellFontInfo(HEADER_FONT_COLOR, null, HEADER_BACKGROUND_COLOR, BOLD_WEIGHT_NORMAL));
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
