package poi.excel.report.api;

import org.apache.poi.ss.usermodel.HorizontalAlignment;

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
    private boolean autofit;
    private HorizontalAlignment alignment = HorizontalAlignment.LEFT;

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

    public boolean isAutofit() {
        return autofit;
    }

    public HorizontalAlignment getAlignment() {
        return alignment;
    }

    public void setAutofit(boolean autofit) {
        this.autofit = autofit;
    }

    public void setAlignment(HorizontalAlignment alignment) {
        this.alignment = alignment;
    }
}
