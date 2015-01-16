package poi.excel.report.api;

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

    public List<ReportField> getRowGroup() {
        return rowGroup;
    }

    public List<ReportField> getReportFields() {
        return reportFields;
    }

    public List<DataField> getDataFields() {
        return dataFields;
    }

    private final List<ReportField> rowGroup;

    public ReportInfo(List<DataField> dataFields, List<ReportField> reportFields, List<ReportField> rowGroup) {
        this.dataFields = dataFields;
        this.reportFields = reportFields;
        this.rowGroup = rowGroup;
    }
}
