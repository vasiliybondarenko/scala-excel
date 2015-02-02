package poi.excel.report;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 2/2/15
 * Time: 12:13 PM
 * Project: scala-excel
 */
public class AutoFilter {
    private final int firstColumn;
    private final int lastColumn;

    public AutoFilter(int firstColumn, int lastColumn) {
        this.firstColumn = firstColumn;
        this.lastColumn = lastColumn;
    }

    public int getFirstColumn() {
        return firstColumn;
    }

    public int getLastColumn() {
        return lastColumn;
    }
}
