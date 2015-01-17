package poi.excel.report.api;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/16/15
 * Time: 3:12 PM
 * Project: scala-excel
 */
public class ReportField {

    private final String field;
    private final String title;
    private String format;

    public ReportField(String field, String title, String format) {
        this(field, title);
        this.format = format;
    }

    public ReportField(String field, String title) {
        this.field = field;
        this.title = title;
    }

    public String getField() {
        return field;
    }

    public String getTitle() {
        return title;
    }

    public String getFormat() {
        return format;
    }
}
