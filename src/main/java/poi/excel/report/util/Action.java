package poi.excel.report.util;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/18/15
 * Time: 9:32 PM
 * Project: scala-excel
 */
public interface Action<T> {
    void apply(T o);
}
