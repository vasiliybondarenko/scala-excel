package poi.excel.report.util;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/18/15
 * Time: 12:01 AM
 * Project: scala-excel
 */
public interface Mapper<V, R> {
    <R> R apply(V value);
}
