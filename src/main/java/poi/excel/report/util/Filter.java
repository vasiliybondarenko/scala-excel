package poi.excel.report.util;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/18/15
 * Time: 12:21 AM
 * Project: scala-excel
 */
public interface Filter<T> {
    boolean apply(T value);
}
