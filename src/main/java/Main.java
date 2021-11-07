import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class Main{
    public static void main(String[] args) throws Exception {

        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("CellSize");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("hi Sergey");
        sheet.setColumnWidth(0, 9000);
        sheet.autoSizeColumn(0);
        row.setHeightInPoints(30);

        sheet.addMergedRegion(new CellRangeAddress(0, 5, 0, 2));




        FileOutputStream fos = new FileOutputStream(
                "y:/JavaAllProjects/poi.apache.org/JavaExelFormulas/CellSize.xlsx");
        wb.write(fos);
        fos.close();
        wb.close();



    }
}
