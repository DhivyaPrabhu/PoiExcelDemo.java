package Apachepoi;


import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class PoiExcelDemo
{
    public static void main(String[] args) throws IOException
    {
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet= workbook.createSheet("Sheet1");
        Object empdata[][]={{"Name","Age","Email"},
                            {"John Doe",30,"john@test.com"},
                            {"Jane Doe",28,"john@test.com"},
                            {"Bob Smith",35,"jacky@example.com"},
                            {"Swapnil",37,"swapnil@example.com"}
                            };
        //Using Normal for loop
        int rows= empdata.length;
        int cols=empdata[0].length;
        System.out.println(rows);
        System.out.println(cols);
        for(int r=0;r<rows;r++)
        {
           XSSFRow row= sheet.createRow(r);
            for (int c=0;c<cols;c++)
            {
                XSSFCell cell=row.createCell(c);
                Object value=empdata[r][c];
                if(value instanceof String)
                    cell.setCellValue((String) value);
                if(value instanceof Integer)
                cell.setCellValue((Integer) value);
                if(value instanceof Boolean)
                    cell.setCellValue((Boolean) value);

            }
        }
        String filePath=".\\Demo\\employee.xlsx";
        FileOutputStream fileOutputStream=new FileOutputStream(filePath);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("Employee.xlsx file written success");
        }
    }
