package poi.excel.report;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

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

    final short TITLE_FONT_COLOR = IndexedColors.GREY_40_PERCENT.getIndex();

    final short TITLE_FONT_SIZE = 20;

    final short HEADER_FONT_COLOR = IndexedColors.WHITE.getIndex();

    final  short HEADER_BACKGROUND_COLOR = IndexedColors.RED.getIndex();
}
