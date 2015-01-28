package poi.excel.report.api;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import poi.excel.report.ReportStyle;

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
    private HorizontalAlignment horizontalAlignment = HorizontalAlignment.LEFT;
    private VerticalAlignment verticalAlignment = VerticalAlignment.BOTTOM;
    private int columnWidth = ReportStyle.DEFAULT_COLUMN_WIDTH;

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

    public HorizontalAlignment getHorizontalAlignment() {
        return horizontalAlignment;
    }

    public void setAutofit(boolean autofit) {
        this.autofit = autofit;
    }

    public void setHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
    }

    public VerticalAlignment getVerticalAlignment() {
        return verticalAlignment;
    }

    public void setVerticalAlignment(VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
    }

    public int getColumnWidth() {
        return columnWidth;
    }

    public void setColumnWidth(int columnWidth) {
        this.columnWidth = columnWidth;
    }
}
