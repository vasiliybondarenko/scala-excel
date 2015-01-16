package poi.excel.report;

import org.apache.poi.xssf.usermodel.XSSFCell;
import poi.excel.report.api.DataType;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/16/15
 * Time: 3:57 PM
 * Project: scala-excel
 */
public class TextConverter extends CellDataConverter{
    public TextConverter(DataType dataType) {
        super(dataType);
    }

    @Override
    public String getDataFormat() {
        return "";
    }

    @Override
    public void assignValue(String value, XSSFCell cell) {
        cell.setCellValue(value);
    }
}
