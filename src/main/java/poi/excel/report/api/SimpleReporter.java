package poi.excel.report.api;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import poi.excel.report.*;
import poi.excel.report.converters.CellDataConverter;
import poi.excel.report.converters.CellDataConverterFactory;
import poi.excel.report.util.ImageHelper;
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

    private final String destinationPath = "/Users/shredinger/Downloads/java_port_test.xlsx";

    //TODO: replace with relative path
    private final String logoImagePath = "src/main/resources/logo.jpg";

    @Override
    public void run(ReportInfo reportInfo, List<List<String>> reportData, List<FooterField> footerFields) {
        final int firstColumnIndex = 0;
        final int firstRowIndex = 5;

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(reportInfo.getSpreadshitName());
        XSSFRow headerRow = sheet.createRow(firstRowIndex - 1);
        List<DataField> dataFields = getFilteredDataFields(reportInfo);
        List<ReportField> reportFields = getFilteredReportFields(reportInfo);
        List<XSSFCellStyle> columnStyles = createColumnStyles(wb, dataFields);

        //title/logo
        final int logoColumn = 0;
        final int titleStartColumnIndex = 2;
        final int titleStartRowIndex = 2;
        final int logoRowIndex = titleStartRowIndex - 1;
        final short logoRowHeight = (short)500;
        final short titleRowHeight = (short)600;
        XSSFRow logoRow = sheet.createRow(logoRowIndex);
        logoRow.setHeight(logoRowHeight);
        logoRow.createCell(logoColumn);
        
        XSSFRow titleRow = sheet.createRow(titleStartRowIndex);
        titleRow.setHeight(titleRowHeight);
        XSSFCell titleCell = titleRow.createCell(titleStartColumnIndex);
        titleCell.setCellValue(reportInfo.getTitleName());
        titleCell.setCellStyle(getTitleRowStyle(wb));
        try {
            //TODO: place image in proper place            
            XSSFCell logoCell = sheet.getRow(logoRowIndex).createCell(logoColumn);
            XSSFCellStyle logoCellStyle = wb.createCellStyle();
            logoCell.setCellStyle(logoCellStyle);            

            ImageHelper.addLogoImageToWorkBook(wb, sheet, logoImagePath, logoColumn, logoRowIndex);

            sheet.autoSizeColumn(logoColumn);
        } catch (Exception e) {
            //TODO: handle exception
            e.printStackTrace();
        }


        //header
        int columnNumber = firstColumnIndex;
        List<ReportField> filteredReportFields = getFilteredReportFields(reportInfo);
        XSSFCellStyle headerRowStyle = getHeaderRowStyle(wb);
        headerRowStyle.setAlignment(CellStyle.ALIGN_CENTER);
        for(ReportField item: filteredReportFields) {
            XSSFCell cell = headerRow.createCell(columnNumber);
            cell.setCellValue(item.getTitle());
            columnNumber = columnNumber + 1;
            cell.setCellStyle(headerRowStyle);
        }
        
        //auto filter
        AutoFilter autoFilter = reportInfo.getAutoFilter();
        sheet.setAutoFilter(
                new CellRangeAddress(
                        firstRowIndex - 1,
                        reportData.size() + firstRowIndex - 1,
                        autoFilter == null ? firstColumnIndex : autoFilter.getFirstColumn(),
                        autoFilter == null ? firstColumnIndex + filteredReportFields.size() - 1 : autoFilter.getLastColumn())
        );
               
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
                XSSFCellStyle style = columnStyles.get(col);
                XSSFDataFormat format = wb.createDataFormat();
                String formatString = columnFormatString != null ? columnFormatString : cellDataConverter.getDataFormat();
                style.setDataFormat(format.getFormat(formatString));
                style.setAlignment(reportFields.get(col).getHorizontalAlignment());
                style.setVerticalAlignment(reportFields.get(col).getVerticalAlignment());
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
        
        //footer
        int firstFooterRow = firstRowIndex + reportData.size() + 2;
        createFooter(wb, sheet, footerFields, firstFooterRow);

        //auto fit/size columns
        columnNumber = firstColumnIndex;
        for(ReportField item: filteredReportFields) {            
            if(item.isAutofit()){
                HSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
                sheet.autoSizeColumn(columnNumber);
            } else {
                sheet.setColumnWidth(columnNumber, item.getColumnWidth());
            }
            columnNumber ++;
        }

        FileOutputStream resultFile = null;
        try {
            resultFile = new FileOutputStream(destinationPath);
            wb.write(resultFile);
            resultFile.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }
    
    private void createFooter(XSSFWorkbook wb, XSSFSheet sheet, List<FooterField> footerFields, int firstRow){
        if(footerFields == null){
            return;
        }
        int row = firstRow;
        int keyColumnIndex = 0;
        int valueColumnIndex = 1;
        for(FooterField footerField: footerFields){
            XSSFRow xssfRow = sheet.createRow(row);
            XSSFCell keyCell = xssfRow.createCell(keyColumnIndex);            
            XSSFCell valueCell = xssfRow.createCell(valueColumnIndex);            
            keyCell.setCellValue(footerField.getLabel());
            valueCell.setCellValue(footerField.getValue());

            CellDataConverter cellDataConverter = CellDataConverterFactory.create(footerField.getDataType());
            cellDataConverter.assignValue(footerField.getValue(), valueCell);

            XSSFCellStyle style = createFooterStyle(wb);
            
            setFooterValueCellDataFormat(wb, style, cellDataConverter);
            keyCell.setCellStyle(style);
            valueCell.setCellStyle(style);
            
            row ++;            
        }        
    }

    private List<XSSFCellStyle> createColumnStyles(XSSFWorkbook wb, List<DataField> dataFields){
        ArrayList<XSSFCellStyle> styles = new ArrayList<>();
        for(DataField dataField: dataFields){
            styles.add(wb.createCellStyle());            
        }
        return styles;        
    }
    
    private XSSFCellStyle createFooterStyle(XSSFWorkbook wb) {
        return getStyle(wb, new CellFontInfo(FOOTER_FONT_COLOR, FOOTER_FONT_SIZE, FOOTER_BG_COLOR, FOOTER_BOLD_WEIGHT));
    }

    private void setFooterValueCellDataFormat(XSSFWorkbook wb, XSSFCellStyle style, CellDataConverter cellDataConverter) {
        String formatString = cellDataConverter.getDataFormat();        
        XSSFDataFormat format = wb.createDataFormat();
        style.setDataFormat(format.getFormat(formatString));        
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
