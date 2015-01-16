package poi.excel.report.api;

import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/16/15
 * Time: 3:29 PM
 * Project: scala-excel
 */
public interface Reporter {
    void run(ReportInfo reportInfo , List<List<String>> reportData);
}
