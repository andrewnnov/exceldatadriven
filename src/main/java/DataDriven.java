import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class DataDriven {

    public static void main(String[] args) throws IOException {

        FileInputStream fileInputStream = new FileInputStream("C:\\java\\excel\\datademo.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        int sheets = workbook.getNumberOfSheets();


        for (int i = 0; i< sheets; i++) {
            if(workbook.getSheetName(i).equalsIgnoreCase("testdata")) {
                XSSFSheet sheet = workbook.getSheetAt(i);
                //System.out.println(sheet.getSheetName());
                Iterator<Row> rows = sheet.iterator(); //sheet is collection of row
                Row firstRow = rows.next();
                Iterator<Cell> cell = firstRow.cellIterator(); //row is collection of cells
                //cell.next();
                int k = 0;
                int column = 0;
                while (cell.hasNext()) {
                    Cell value = cell.next();
                    if(value.getStringCellValue().equalsIgnoreCase("TestCases")) {
                        column = k;
                    }
                    k++;
                }
                System.out.println(column);

                while (rows.hasNext()) {
                    Row r = rows.next();
                    if(r.getCell(column).getStringCellValue().equalsIgnoreCase("Purchase")) {
                        Iterator<Cell> cv = r.cellIterator();
                        while (cv.hasNext()) {

                            System.out.println(cv.next().getStringCellValue());

                        }
                    }
                }
            }
        }
    }
}
