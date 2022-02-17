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
public class InitDataExcel {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        String dataPath = "C:\\Users\\zhao\\Desktop\\commit数据模版.xlsx";
        String dataOutPath = "C:\\Users\\zhao\\Desktop\\04.xlsx";
        File excel = new File(dataPath);
        Workbook wb = new XSSFWorkbook(excel);
        Sheet sheet = wb.getSheetAt(0);     //读取sheet 0
        Sheet sheetpro = wb.getSheetAt(1);
        int firstRowIndex = 32;   //第一行是列名，所以不读
        int lastRowIndex = 69;
        int start = 1;
        for(int rIndex = firstRowIndex ; rIndex <= lastRowIndex; rIndex++) {   //遍历行
            //System.out.println("rIndex: " + rIndex);
            Row row = sheet.getRow(rIndex);
            Cell numberCell = row.getCell(2);
            Cell comIdCell = row.getCell(3);
            for(int i = start; i <= 14176; i ++){
                Row rowInit = sheetpro.getRow(i);
                Cell cellProNum = rowInit.createCell(0);
                cellProNum.setCellValue(6);
                Cell cellProName = rowInit.createCell(1);
                cellProName.setCellValue("mogu_blog");
                Cell initComCell = rowInit.getCell(5);
                Cell cellNum = rowInit.createCell(2);
                cellNum.setCellValue((int)(Double.parseDouble(numberCell.toString())));
                Cell cellName = rowInit.createCell(3);
                cellName.setCellValue(comIdCell.toString());
                if(initComCell.toString().equals(comIdCell.toString()) == false &&
                        sheetpro.getRow(i - 1).getCell(5).toString().equals(comIdCell.toString())){
                   start = i;
                    break;
                }
            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream(dataOutPath);
        wb.write(fileOutputStream);
        fileOutputStream.flush();
        fileOutputStream.close();
    }


}
