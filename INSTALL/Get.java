
import org.apache.poi.xssf.usermodel.*;
import java.io.FileInputStream;
import java.io.*;
import java.util.*;
import java.lang.*;
public class Get {


    public static void main(String[] args)throws Exception {
        Scanner sc=new Scanner(System.in);
        InputStream inp = new FileInputStream(sc.next());
        XSSFWorkbook workbook = new XSSFWorkbook(inp);//读取现有的Excel文件
        List<XSSFPictureData> pictures = workbook.getAllPictures();

            if (pictures.size()!=0)//判断文件格式
            {
                System.out.println("该Excel有图片");
            }
            if(pictures.size() ==0)
            {
                System.out.println("该Excel无图片");
            }


    }
}