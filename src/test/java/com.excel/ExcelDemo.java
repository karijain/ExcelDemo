package com.excel;

import com.sun.xml.internal.fastinfoset.tools.XML_SAX_StAX_FI;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

public class ExcelDemo {

    public static void main(String args[]) throws Exception {
        FileInputStream inputStream=new FileInputStream(new File("src\\test\\resources\\demo.xls"));
        Workbook workbook=new HSSFWorkbook(inputStream);
        Sheet sheet=workbook.getSheet("Sheet1");
        int rowCount=sheet.getLastRowNum()-sheet.getFirstRowNum();

        for(int i=0;i<rowCount+1;i++)
        {
            Row row=sheet.getRow(i);
            for(int j=0;j<row.getLastCellNum();j++)
            {
               // System.out.println(row.getCell(i).getCellType());
                String value=row.getCell(j).getStringCellValue();
                System.out.println(value+ " ");
            }
        }

        Row row=sheet.getRow(0);
        String[] arr={"F","F"};
        Row newRow=sheet.createRow(rowCount+1);


        // rowCount=sheet.getLastRowNum()-sheet.getFirstRowNum();

        for(int i=0;i<row.getLastCellNum();i++)
        {
            newRow.createCell(i).setCellValue(arr[i]);
        }

        inputStream.close();

        FileOutputStream outputStream=new FileOutputStream(new File("src\\test\\resources\\demo.xls"));
        workbook.write(outputStream);
        outputStream.close();


 }

}
