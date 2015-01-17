package poi.excel.report;

import org.apache.poi.xssf.usermodel.XSSFCell;
import poi.excel.report.api.DataType;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/16/15
 * Time: 2:32 PM
 * Project: scala-excel
 */
public abstract class CellDataConverter{
    final DataType dataType;

    public CellDataConverter(DataType dataType) {
        this.dataType = dataType;
    }

    public abstract String getDataFormat();
    public abstract void assignValue(String value, XSSFCell cell);
}
