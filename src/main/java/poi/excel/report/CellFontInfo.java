package poi.excel.report;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/20/15
 * Time: 5:36 PM
 * Project: scala-excel
 */
public class CellFontInfo {

    public final Short fontColor;
    public final  Short fontSize;
    public final  Short backgroundColor;
    public final  Short boldWeight;

    public CellFontInfo(Short fontColor, Short fontSize, Short backgroundColor, Short boldWeight) {
        this.fontColor = fontColor;
        this.fontSize = fontSize;
        this.backgroundColor = backgroundColor;
        this.boldWeight = boldWeight;
    }
}
