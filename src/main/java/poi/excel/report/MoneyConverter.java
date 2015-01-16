package poi.excel.report;

import org.apache.poi.xssf.usermodel.XSSFCell;
import poi.excel.report.api.DataType;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/16/15
 * Time: 3:47 PM
 * Project: scala-excel
 */
public class MoneyConverter extends CellDataConverter{

    private final String DEFAULT_CURRENCY_FORMAT = "$#,##0.00";

    public MoneyConverter(DataType dataType) {
        super(dataType);
    }

    @Override
    public String getDataFormat() {
        return DEFAULT_CURRENCY_FORMAT;
    }

    @Override
    public void assignValue(String value, XSSFCell cell) {
        cell.setCellValue(Double.parseDouble(value));
    }
}
