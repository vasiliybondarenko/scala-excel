package poi.excel.report.converters;


import poi.excel.report.api.DataType;

import java.util.HashMap;
import java.util.Map;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/16/15
 * Time: 4:00 PM
 * Project: scala-excel
 */
public class CellDataConverterFactory {

    private final static Map<DataType, CellDataConverter> converters = new HashMap<>();

    static {
        converters.put(DataType.Number, new NumberConverter(DataType.Number));
        converters.put(DataType.Money, new MoneyConverter(DataType.Money));
        converters.put(DataType.Date, new DateConverter(DataType.Date));
        converters.put(DataType.String, new TextConverter(DataType.String));
    }

    public static CellDataConverter create(DataType dataType){
        return converters.get(dataType);
    }
}
