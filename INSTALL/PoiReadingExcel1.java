package test;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Test;

import java.util.Scanner;
import java.io.FileInputStream;


public class PoiReadingExcel1 {
    public static void main(String[] args) throws Exception {
        Scanner sc=new Scanner(System.in);
        testNumber();
        //sheet的第i个数据
        System.out.println("请输入要获取sheet名称的序号：");
        int i=sc.nextInt();
        //第j行
        System.out.println("请输入要获取数据的行数：");
        int j=sc.nextInt();
        //第j行的第k个数据
        System.out.println("请输入要获取数据的列数：");
        int k=sc.nextInt();
        testRead(i,j,k);
    }
    @Test
    public static void testNumber() throws Exception {
        //读取文件的位置
        String path = "/Users/wp/Desktop/test.xlsx";
        //获取文件输入流
        FileInputStream fileInputStream = new FileInputStream(path);
        //通过文件流创建（获取）工作簿,excel中的操作，Java基本都能实现，这里的新建的对象注意与excel版本对应
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        //获取工作表sheet,0表示获取第一个sheet
        Sheet sheet = workbook.getSheetAt(0);
        //获得sheet
        Row rowHead = sheet.getRow(0);
        //获得sheet的总个数
        int sheetNumber=rowHead.getPhysicalNumberOfCells();
        System.out.println("该Excel共有"+sheetNumber+"个sheet");
        //获得数据的总行数
        int totalRowNum = sheet.getLastRowNum();
        System.out.println("该Excel共有"+totalRowNum+"行");
    }

    @Test
    public static void testRead(int i,int j,int k) throws Exception {
        //读取文件的位置
        String path = "/Users/wp/Desktop/test.xlsx";
        //获取文件输入流
        FileInputStream fileInputStream = new FileInputStream(path);
        //通过文件流创建（获取）工作簿,excel中的操作，Java基本都能实现，这里的新建的对象注意与excel版本对应
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        //获取工作表sheet,0表示获取第一个sheet
        Sheet sheet = workbook.getSheetAt(0);
        //获得sheet
        Row rowHead = sheet.getRow(0);
        //获得sheet的总个数
        int sheetNumber=rowHead.getPhysicalNumberOfCells();
        //获得数据的总行数
        int totalRowNum = sheet.getLastRowNum();
        //获取sheet行第i个数据
        if(i<=totalRowNum) {
            Cell cell = rowHead.getCell(i-1);
            System.out.println("该excel中第" + i + "个sheet的名字为：" + cell.getStringCellValue());
        }
        else{
            System.out.println("该Excel没有第"+i+"个sheet");
        }
        //获取第j行数据
        if(j>=1&&j<=totalRowNum) {
            //获取第j行数据
            Row row = sheet.getRow(j);
            if(k<=sheetNumber) {
                //获取第j行第k个数据
                Cell cell1 = row.getCell(k-1);
                System.out.println("该Excel第" + j + "行第" + k + "个数据为：" + cell1.getStringCellValue());
            }
            else {
                System.out.println("该Excel没有第"+k+"列");
            }
        }
        else{
            System.out.println("该Excel中没有第"+j+"行");
        }
    }
}