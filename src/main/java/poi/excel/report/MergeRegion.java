package poi.excel.report;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/18/15
 * Time: 9:24 PM
 * Project: scala-excel
 */
public class MergeRegion {
    private final String field;
    private final int firstRow;
    private final int lastRow;

    public MergeRegion(String field, int firstRow, int lastRow) {
        this.field = field;
        this.firstRow = firstRow;
        this.lastRow = lastRow;
    }

    public String getField() {
        return field;
    }

    public int getFirstRow() {
        return firstRow;
    }

    public int getLastRow() {
        return lastRow;
    }
}
