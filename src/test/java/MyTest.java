import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;
import java.util.LinkedList;
import java.util.List;

/**
 * Created by phoenix on 2/28/17.
 */
public class MyTest {
    String filename = "test";
    String path = "src/test/resources";
    String typei = "xlsx";
    String typeo = "xlsx";
    String filenameIn = path+"/in/"+filename+"."+typei;
    String filenameOut = path+"/out/"+filename+"."+typeo;

    private static String[] daysOfWeek = {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"};

    @Test
    public void test() throws Exception {
        // create workbook
        XSSFWorkbook wb = new XSSFWorkbook();
        // create sheets
        List<Sheet> sheets = new LinkedList<>();
        sheets.add(wb.createSheet("jan-17"));
        sheets.add(wb.createSheet("feb-17"));

        // process each created sheet
        sheets.forEach(sheet->{
            // create a header row
            Row row = sheet.createRow(0);
            // add header data
            for (int i = 0; i < daysOfWeek.length; i++) {
                Cell cell = row.createCell(i, Cell.CELL_TYPE_STRING);
                cell.setCellValue(daysOfWeek[i]);
            }
        });
        // write data out
        // InputStream ins = new FileInputStream(filenameIn);
        FileOutputStream outs = new FileOutputStream(filenameOut);

        wb.write(outs);

        outs.close();
    }
}
