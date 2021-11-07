import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;

public class Main {
    public static void main(String[] args) throws Exception{
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet1");

        Row row = sheet.createRow(0);
        Cell cell0 = row.createCell(0);
        cell0.setCellValue("Hi Sergey");

        CellStyle style = wb.createCellStyle();
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        //style.setFillBackgroundColor(IndexedColors.GREEN.getIndex());
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_TOP);
        style.setBorderBottom(CellStyle.BORDER_DASH_DOT_DOT);
        style.setBottomBorderColor(IndexedColors.GREEN.getIndex());

        Font font = wb.createFont();
        font.setFontName("Courier New");
        font.setFontHeightInPoints((short) 15);
        font.setBold(true);
        font.setStrikeout(true);
        font.setUnderline(Font.U_SINGLE);
        font.setColor(IndexedColors.RED.getIndex());

        style.setFont(font);

        cell0.setCellStyle(style);

        FileOutputStream fos = new FileOutputStream(
                "y:/JavaAllProjects/poi.apache.org/JavaExelFormulas/Styles.xls");
        wb.write(fos);
        fos.close();
        wb.close();
    }
}

