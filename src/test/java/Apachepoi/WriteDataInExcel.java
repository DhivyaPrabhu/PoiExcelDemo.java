package Apachepoi;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class WriteDataInExcel
{
    public static void main(String[] args) throws IOException
    {
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet= workbook.createSheet("StudentDetails");
        Object studentdata[][]={{"Name","Age","Email"},
                {"Arun",20,"arun@test.com"},
                {"Varun",19,"varun@test.com"},
                {"Sri Gupta",18,"sri@example.com"},
                {"Swapna",17,"swapna@example.com"}
        };
        //Using Normal for loop
        int rows= studentdata.length;
        int cols=studentdata[0].length;

        for(int r=0;r<rows;r++)
        {
            XSSFRow row= sheet.createRow(r);
            for (int c=0;c<cols;c++)
            {
                XSSFCell cell=row.createCell(c);
                Object value=studentdata[r][c];
                if(value instanceof String)
                    cell.setCellValue((String) value);
                if(value instanceof Integer)
                    cell.setCellValue((Integer) value);
                if(value instanceof Boolean)
                    cell.setCellValue((Boolean) value);

            }
        }
        String filePath=".\\Demo\\StudentsBio.xlsx";
        FileOutputStream fileOutputStream=new FileOutputStream(filePath);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("Employee.xlsx file written success");
    }
}
