package poi.excel.report.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/20/15
 * Time: 9:36 PM
 * Project: scala-excel
 */
public class ImageHelper {
    public static void addLogoImageToWorkBook(XSSFWorkbook wb, Sheet sheet, String imageFilePath, int column, int row) throws Exception{
        byte[] bytes = readImageAsBytes(imageFilePath);
        int pictureIdx = wb.addPicture(bytes, wb.PICTURE_TYPE_JPEG);

        CreationHelper helper = wb.getCreationHelper();

        Drawing drawing = sheet.createDrawingPatriarch();

        ClientAnchor anchor = helper.createClientAnchor();

        anchor.setCol1(column);
        anchor.setRow1(row);
        Picture pict = drawing.createPicture(anchor, pictureIdx);

        pict.resize();

    }

    private static byte[] readImageAsBytes(String imageFilePath) throws IOException {
        InputStream is = new FileInputStream(imageFilePath);
        try {
            return IOUtils.toByteArray(is);
        }finally {
            is.close();
        }
    }
}
