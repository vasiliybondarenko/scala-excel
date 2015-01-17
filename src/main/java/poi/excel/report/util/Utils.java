package poi.excel.report.util;

import java.util.HashSet;
import java.util.List;
import java.util.Set;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/16/15
 * Time: 5:05 PM
 * Project: scala-excel
 */
public class Utils {
    public static <T> Set<T> toSet(List<T> l){
        HashSet<T> set = new HashSet<>();
        for(T e: l){
            set.add(e);
        }
        return set;
    }


}
