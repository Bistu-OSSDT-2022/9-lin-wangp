package test;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;


import java.io.FileInputStream;

public class texs {
    @Test
    public void testRead() throws Exception {
        //读取文件的位置
        String path = "/Users/wp/Desktop/test.xlsx";
        //获取文件输入流
        FileInputStream fileInputStream = new FileInputStream(path);
        //通过文件流创建（获取）工作簿,excel中的操作，Java基本都能实现，这里的新建的对象注意与excel版本对应
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        //获取工作表sheet,0表示获取第一个sheet
        Sheet sheet = workbook.getSheetAt(0);
        //获取第一行数据
        Row row = sheet.getRow(1);
        //获取第一行第一个数据
        Cell cell = row.getCell(0);
        System.out.println(cell.getStringCellValue());
    }
}



