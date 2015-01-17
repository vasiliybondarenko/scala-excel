package poi.excel.report.api;

import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/16/15
 * Time: 3:27 PM
 * Project: scala-excel
 */
public class ExcelReporter implements Reporter{

    private final Reporter reporter = new SimpleReporter();

    @Override
    public void run(ReportInfo reportInfo, List<List<String>> reportData) {
        reporter.run(reportInfo, reportData);
    }
}
