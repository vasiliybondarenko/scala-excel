package poi.excel.report.converters;

import org.apache.poi.xssf.usermodel.XSSFCell;
import poi.excel.report.api.DataType;

import java.text.ParseException;
import java.text.SimpleDateFormat;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/16/15
 * Time: 3:50 PM
 * Project: scala-excel
 */
public class DateConverter extends CellDataConverter{

    public DateConverter(DataType dataType) {
        super(dataType);
    }

    @Override
    public String getDataFormat() {
        return "yyyy/mm/dd";
    }

    @Override
    public void assignValue(String value, XSSFCell cell) {
        SimpleDateFormat formatter = new SimpleDateFormat("yyyy/mm/dd");
        try {
            cell.setCellValue(formatter.parse(value));
        } catch (ParseException e) {
            throw new IllegalArgumentException(e);
        }
    }
}
