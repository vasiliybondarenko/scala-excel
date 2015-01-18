package poi.excel.report.util;

import java.util.*;

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

    public static <T1, T2> List<Tuple<T1, T2>> zip(List<T1> l1, List<T2> l2){
        int size = Math.min(l1.size(), l2.size());
        ArrayList<Tuple<T1, T2>> tuples = new ArrayList<>();
        for(int i = 0; i < size; i ++){
            tuples.add(new Tuple<T1, T2>(l1.get(i), l2.get(i)));
        }
        return tuples;
    }

    public static <V, R> List<R> map(List<V> source, Mapper<V, R> mapper){
        ArrayList<R> resultList = new ArrayList<>();
        for(V item: source){
            resultList.add((R)mapper.apply(item));
        }
        return resultList;
    }

    public static <T> List<T> filter(List<T> source, Filter<T> predicate){
        ArrayList<T> filtered = new ArrayList<>();
        for (T item: source){
            if(predicate.apply(item)){
                filtered.add(item);
            }
        }
        return filtered;
    }

    public static <T1, T2> Map<T1, T2> toMap(List<Tuple<T1, T2>> l){
        HashMap<T1, T2> map = new HashMap<>();
        for(Tuple<T1, T2> tuple: l){
            map.put(tuple.filed1, tuple.filed2);
        }
        return map;
    }


    public static class Tuple<T1, T2>{
        private final T1 filed1;
        private final T2 filed2;

        public Tuple(T1 filed1, T2 filed2) {
            this.filed1 = filed1;
            this.filed2 = filed2;
        }

        public T1 getFiled1() {
            return filed1;
        }

        public T2 getFiled2() {
            return filed2;
        }
    }

}
