import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;

public class Main {
    public static void main(String[] args) throws Exception{
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("Formulas");

        Row row = sheet.createRow(0);

        Cell cell0 = row.createCell(0);
        cell0.setCellValue(2);

        Cell cell1 = row.createCell(1);
        cell1.setCellValue(7);

        Cell cell2 = row.createCell(2);
        cell2.setCellFormula("A1+B1");

        Row row10 = sheet.createRow(10);
        Cell cell10 = row10.createCell(0);
        cell10.setCellValue(5);
        Row row11 = sheet.createRow(11);
        Cell cell11 = row11.createCell(0);
        cell11.setCellValue(6);
        Row row12 = sheet.createRow(12);
        Cell cell12 = row12.createCell(0);
        cell12.setCellValue(7);
        Row row13 = sheet.createRow(13);
        Cell cell13 = row13.createCell(0);
        cell13.setCellValue(8);

        Row row14 = sheet.createRow(14);
        Cell cell14 = row14.createCell(0);
        cell14.setCellFormula("SUM(A11:A14)");



        FileOutputStream fos = new FileOutputStream(
                "y:/JavaAllProjects/poi.apache.org/JavaExelFormulas/Formula.xls");
        wb.write(fos);
        fos.close();
        wb.close();
    }
}
