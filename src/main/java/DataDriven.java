import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class DataDriven {

    public static void main(String[] args) throws IOException {


    }


    public ArrayList<String> getData(String testCaseName) throws IOException {
        ArrayList<String> arrayList = new ArrayList<String>();

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
                    if(r.getCell(column).getStringCellValue().equalsIgnoreCase(testCaseName)) {
                        Iterator<Cell> cv = r.cellIterator();
                        while (cv.hasNext()) {
                            Cell c = cv.next();
                            if(c.getCellTypeEnum() == CellType.STRING) {
                                arrayList.add(c.getStringCellValue());
                            }else {

                                arrayList.add(NumberToTextConverter.toText(c.getNumericCellValue()));
                            }


                        }
                    }
                }
            }
        }

        return arrayList;

    }
}
