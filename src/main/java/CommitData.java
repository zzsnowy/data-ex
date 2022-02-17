import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import static java.lang.Math.abs;
/**
 * @author zzsnowy
 * @date 2022/2/17
 */
public class CommitData {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        String dataPath = "C:\\Users\\zhao\\Desktop\\06-mogu_blog.xlsx";
        File excel = new File(dataPath);
        Workbook wb = new XSSFWorkbook(excel);

        Sheet sheet = wb.getSheetAt(1);     //读取sheet 0

        int firstRowIndex = sheet.getFirstRowNum();   //第一行是列名，所以不读
        int lastRowIndex = sheet.getLastRowNum();

        for(int rIndex = firstRowIndex ; rIndex <= lastRowIndex; rIndex++) {   //遍历行
            //System.out.println("rIndex: " + rIndex);
            Row row = sheet.getRow(rIndex);
            Cell numberCell = row.getCell(2);
            Cell comIdCell = row.getCell(3);
//            if(numberCell != null){
//            System.out.println((int)(Double.parseDouble(numberCell.toString())));
//            }
            if(comIdCell != null){
                System.out.println(comIdCell.toString());
            }
        }

    }


}
