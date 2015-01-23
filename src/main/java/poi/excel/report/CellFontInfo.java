package poi.excel.report;

import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/20/15
 * Time: 5:36 PM
 * Project: scala-excel
 */
public class CellFontInfo {

    public final XSSFColor fontColor;
    public final  Short fontSize;
    public final  XSSFColor backgroundColor;
    public final  Short boldWeight;

    public CellFontInfo(XSSFColor fontColor, Short fontSize, XSSFColor backgroundColor, Short boldWeight) {
        this.fontColor = fontColor;
        this.fontSize = fontSize;
        this.backgroundColor = backgroundColor;
        this.boldWeight = boldWeight;
    }
}
