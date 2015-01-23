package poi.excel.report;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.awt.*;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/20/15
 * Time: 5:50 PM
 * Project: scala-excel
 */
public interface ReportStyle {
    final short BOLD_WEIGHT = Font.BOLDWEIGHT_BOLD;

    final short BOLD_WEIGHT_NORMAL = Font.BOLDWEIGHT_NORMAL;

    final XSSFColor TITLE_FONT_COLOR = new XSSFColor(Color.lightGray);//IndexedColors.GREY_40_PERCENT.getIndex();

    final short TITLE_FONT_SIZE = 20;

    final XSSFColor HEADER_FONT_COLOR = new XSSFColor(Color.white);

    final  XSSFColor HEADER_BACKGROUND_COLOR = new XSSFColor(Color.red);

    final XSSFColor FOOTER_FONT_COLOR = new XSSFColor(Color.red);

    final short FOOTER_FONT_SIZE = 10;

    final short FOOTER_BOLD_WEIGHT = Font.BOLDWEIGHT_BOLD;
    
    final XSSFColor FOOTER_BG_COLOR = new XSSFColor(new Color(255, 255, 255));

}
