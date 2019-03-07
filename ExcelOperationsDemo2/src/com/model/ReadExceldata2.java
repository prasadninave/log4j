package com.model;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.model.ReadExceldata2;

public class ReadExceldata2 {
	int rownum=0;
	int column=0;
	
	public void readExcel(String file, String sheetname) throws IOException
	{
		//int arrayexceldata[][]= null;
		
		FileInputStream fis= new FileInputStream(file);
	
		XSSFWorkbook wb= new XSSFWorkbook(fis);
		XSSFSheet sheet= wb.getSheet(sheetname);
		XSSFRow row= sheet.getRow(3);
		XSSFCell cell= row.getCell(2);
		String val1= cell.getStringCellValue();
		System.out.println("The String at index 2,2 is:"+val1);
		
		int rows= sheet.getLastRowNum();
		int rowcount= rows+1;
		System.out.println("The number of row are:"+rowcount);
		
		int columns= sheet.getRow(rows).getLastCellNum();
		System.out.println("The numbers of coloums are:"+columns);
		
		int arrayexceldata[][]  = new int[rowcount][columns];
		for(int i=0; i<rowcount;i++)
		{
			for(int j=0; j<columns;j++)
			{
				//System.out.println(sheet.getRow(i).getCell(j));
				
				DataFormatter dataformat= new DataFormatter();
				
				String val= dataformat.formatCellValue((sheet.getRow(i).getCell(j)));
				
				System.out.println(val);
				
				WriteExceldata2 obj1 = new WriteExceldata2();
				
				obj1.setcelldata("E:\\Workspace\\ExcelOperationsDemo2\\datawrite.xlsx", "Sheet1", rownum++, column, val);
				
			}
		}
		
		
		
		
	}

	public static void main(String[] args) throws IOException {
		ReadExceldata2 obj = new ReadExceldata2();
		
		obj.readExcel("E:\\Workspace\\ExcelOperationsDemo2\\ExcelOperations.xlsx","Sheet1");
		
		//WriteExceldata2 obj1 = new WriteExceldata2();
		//obj1.setcelldata("E:\\Selenium\\WriteExcelOperations.xlsx","Sheet1", rownum, "column", "s");
		
		
		
	}
}
