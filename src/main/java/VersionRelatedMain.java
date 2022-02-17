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
public class VersionRelatedMain {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        String dataPath = "C:\\Users\\zhao\\Desktop\\06-mogu_blog.xlsx";
        String dataOutPath = "C:\\Users\\zhao\\Desktop\\04.xlsx";
        File excel = new File(dataPath);
        Workbook wb = new XSSFWorkbook(excel);

        Sheet sheet = wb.getSheetAt(1);     //读取sheet 0

        int firstRowIndex = sheet.getFirstRowNum();   //第一行是列名，所以不读
        int lastRowIndex = sheet.getLastRowNum();
        //System.out.println("firstRowIndex: "+firstRowIndex);
        //System.out.println("lastRowIndex: "+lastRowIndex);
        int sum = 3;
        int day = 18;
        int mon = 7;
        int start = 0;
        for(int rIndex = firstRowIndex + 1; rIndex <= lastRowIndex; rIndex++) {   //遍历行
            //System.out.println("rIndex: " + rIndex);
            Row row = sheet.getRow(rIndex);
            if (row != null) {
                Cell cellTime = row.getCell(0);
                Cell cellCommit = row.getCell(1);
                int com = (int)(Double.parseDouble(cellCommit.toString()));
                String date = cellTime.toString();
                //System.out.println(cellTime.toString() + " " + cellCommit.toString());
                int daytemp = Integer.valueOf(date.split("-")[0]);
                String smon = date.split("-")[1];
                int montemp = trans(smon);
                if((montemp != mon && daytemp >= day) ||((montemp - mon) >= 2) ||
                        ((montemp - mon) >= (-10)) && (montemp - mon) < 0){
                    Cell cellComSum = sheet.getRow(start).createCell(4);
                    cellComSum.setCellValue(sum);
                    String lastT = sheet.getRow(rIndex - 1).getCell(0).toString();
                    String lastmon = lastT.split("-")[1];
                    int lastMon = trans(lastmon);
                    String commit = find(wb.getSheetAt(0), Integer.parseInt(lastT.split("-")[2]),
                            lastMon, Integer.parseInt(lastT.split("-")[0]));
                    //System.out.println(commit);
                    Cell cellCom = sheet.getRow(rIndex - 1).createCell(5);
                    cellCom.setCellValue(commit);
                    sum = com;
                    day = daytemp;
                    mon = montemp;
                    start = rIndex;

                }else{
                    sum += com;
                }


            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream(dataOutPath);
        wb.write(fileOutputStream);
        fileOutputStream.flush();
        fileOutputStream.close();


    }

    private static String find(Sheet sheetAt, int year, int mon, int day) {

        String commit = "";
        int n,y,r;
        for(int rIndex = 0; rIndex <= 1633; rIndex++){
            Row row = sheetAt.getRow(rIndex);
            if (row != null) {
                Cell cell = row.getCell(4);
                if(cell.toString().split("/").length != 3){
                    continue;
                }
                //double d = Double.parseDouble(cell.toString().split("/")[0]);
                //System.out.println(cell.toString().split("/")[0] + "  "+ d);
                n = Integer.parseInt(cell.toString().split("/")[0]);
                y = Integer.parseInt(cell.toString().split("/")[1]);
                r = Integer.parseInt(cell.toString().split("/")[2]);
                if(year == n && mon == y && day == r){
                    commit = row.getCell(1).toString();
                    break;
                }
            }
        }
        return commit;

    }


    public static int trans(String smon) {
        int mon = 0;
        if(smon.equals("一月")) mon = 1;
        if(smon.equals("二月")) mon = 2;
        if(smon.equals("三月")) mon = 3;
        if(smon.equals("四月")) mon = 4;
        if(smon.equals("五月")) mon = 5;
        if(smon.equals("六月")) mon = 6;
        if(smon.equals("七月")) mon = 7;
        if(smon.equals("八月")) mon = 8;
        if(smon.equals("九月")) mon = 9;
        if(smon.equals("十月")) mon = 10;
        if(smon.equals("十一月")) mon = 11;
        if(smon.equals("十二月")) mon = 12;
        return mon;
    }
}
