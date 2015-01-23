package poi.excel.report;

import poi.excel.report.api.DataType;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/23/15
 * Time: 4:25 PM
 * Project: scala-excel
 */
public class FooterField {
    private final String label;
    private final String value;
    private final DataType dataType;

    public String getLabel() {
        return label;
    }

    public String getValue() {
        return value;
    }

    public DataType getDataType() {
        return dataType;
    }

    public FooterField(String label, String value, DataType dataType) {
        this.label = label;
        this.value = value;
        this.dataType = dataType;
    }
}
