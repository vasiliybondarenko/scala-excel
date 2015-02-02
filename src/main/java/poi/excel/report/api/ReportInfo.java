package poi.excel.report.api;

import poi.excel.report.AutoFilter;

import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/16/15
 * Time: 3:20 PM
 * Project: scala-excel
 */
public class ReportInfo {
    private final List<DataField> dataFields;
    private final List<ReportField> reportFields;
    private final List<ReportField> rowGroup;
    private final String titleName;
    private final String spreadshitName;   

    private AutoFilter autoFilter;

    public List<ReportField> getRowGroup() {
        return rowGroup;
    }

    public List<ReportField> getReportFields() {
        return reportFields;
    }

    public List<DataField> getDataFields() {
        return dataFields;
    }

    public String getTitleName() {
        return titleName;
    }

    public String getSpreadshitName() {
        return spreadshitName;
    }

    public void setAutoFilter(AutoFilter autoFilter) {
        this.autoFilter = autoFilter;
    }

    public AutoFilter getAutoFilter() {
        return autoFilter;
    }
    
    public ReportInfo(List<DataField> dataFields, List<ReportField> reportFields, List<ReportField> rowGroup, String titleName, String spreadshitName) {
        this.dataFields = dataFields;
        this.reportFields = reportFields;
        this.rowGroup = rowGroup;
        this.titleName = titleName;
        this.spreadshitName = spreadshitName == null ? "Title name" : spreadshitName;
    }
}
