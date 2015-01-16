package poi.excel.report.api;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import poi.excel.report.util.Utils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;

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

    private List<DataField> getFilteredDataFields(ReportInfo reportInfo){
        Set<ReportField> reportFields = Utils.toSet(reportInfo.getReportFields());
        ArrayList<DataField> dataFields = new ArrayList<>();
        for(DataField dataField: reportInfo.getDataFields()){
            if(reportFields.contains(dataField)){
                dataFields.add(dataField);
            }
        }
        return dataFields;
    }

    private List<ReportField> getFilteredReportFields(ReportInfo reportInfo){
        return reportInfo.getReportFields();
    }
}
