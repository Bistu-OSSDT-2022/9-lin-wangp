package test;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Test;

import java.io.InputStream;
import java.util.List;
import java.util.Scanner;
import java.io.FileInputStream;


public class PoiReadingExcel2 {
    public static void main(String[] args) throws Exception {
        Scanner sc=new Scanner(System.in);
        System.out.println("请输入文件所在地址：");
        String d = sc.next();
        tp(d);
        testNumber(d);
        testRead(d);
    }
    @Test
    public static void testNumber(String d) throws Exception {
        //读取文件的位置
        String path = d;
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
    public static void testRead(String d) throws Exception {
        Scanner sc = new Scanner(System.in);
        //读取文件的位置
        String path = d;
        //获取文件输入流
        FileInputStream fileInputStream = new FileInputStream(path);
        //通过文件流创建（获取）工作簿,excel中的操作，Java基本都能实现，这里的新建的对象注意与excel版本对应
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        //获取工作表sheet,0表示获取第一个sheet
        Sheet sheet = workbook.getSheetAt(0);
        //获得sheet
        Row rowHead = sheet.getRow(0);
        //获得sheet的总个数
        int sheetNumber = rowHead.getPhysicalNumberOfCells();
        //获得数据的总行数
        int totalRowNum = sheet.getLastRowNum();
        //获取sheet行第i个数据
        System.out.println("是否需要获取sheet的名称（y/n）:");
        char s = sc.next().charAt(0);
        if (s == 'y') {
            System.out.println("请输入要获取sheet名称的序号:");
            int i = sc.nextInt();
            if (i <= totalRowNum) {
                Cell cell = rowHead.getCell(i - 1);
                System.out.println("该excel中第" + i + "个sheet的名字为：" + cell.getStringCellValue());
            } else {
                System.out.println("该Excel没有第" + i + "个sheet");
            }
        }

        System.out.println("是否需要获取一整行的数据（y/n）:");
        char h = sc.next().charAt(0);
        if (h == 'y') {
            System.out.println("请输入要获取数据的行数：");
            int j = sc.nextInt();
            if (j >= 1 && j <= totalRowNum) {
                //获取第j行数据
                Row row = sheet.getRow(j);
                for (int k = 0; k <= sheetNumber-1; k++) {
                    Cell cell1 = row.getCell(k);
                    int a = k + 1;
                    System.out.println("第" + j + "行第" + a + "个数据为" + cell1.getStringCellValue());
                }
            }
        }

        System.out.println("是否需要获取某一行某一列的具体数据（y/n）:");
        char t = sc.next().charAt(0);
        if (t == 'y') {
            System.out.println("请输入要获取数据的行数：");
            //第j行
            int j=sc.nextInt();
            //第j行的第k个数据
            System.out.println("请输入要获取数据的列数：");
            int k=sc.nextInt();
            //获取第j行数据
            if (j >= 1 && j <= totalRowNum) {
                //获取第j行数据
                Row row1 = sheet.getRow(j);
                if (k <= sheetNumber) {
                    //获取第j行第k个数据
                    Cell cell2 = row1.getCell(k - 1);
                    System.out.println("该Excel第" + j + "行第" + k + "个数据为：" + cell2.getStringCellValue());
                } else {
                    System.out.println("该Excel没有第" + k + "列");
                }
            } else {
                System.out.println("该Excel中没有第" + j + "行");
            }
        }
    }

    public static void tp(String d)throws Exception {

        InputStream inp = new FileInputStream(d);
        XSSFWorkbook workbook = new XSSFWorkbook(inp);//读取现有的Excel文件
        List<XSSFPictureData> pictures = workbook.getAllPictures();

        if (pictures.size()!=0)//判断文件格式
        {
            System.out.println("该Excel有.png图片");
            return;
        }
        if(pictures.size() ==0)
        {
            System.out.println("该Excel无.png图片");
        }
    }
}