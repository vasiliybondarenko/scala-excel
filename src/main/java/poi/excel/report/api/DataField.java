package poi.excel.report.api;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/16/15
 * Time: 3:09 PM
 * Project: scala-excel
 */
public class DataField {
    private final String field;
    private final DataType dataType;

    public DataField(String field, DataType dataType) {
        this.field = field;
        this.dataType = dataType;
    }

    public String getField() {
        return field;
    }

    public DataType getDataType() {
        return dataType;
    }
}
