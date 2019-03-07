package com.model;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExceldata2 {
	
	public void setcelldata(String file,String sheetnm,int rowno,int colno,String dataval) throws IOException
	{
		FileInputStream fis1= new FileInputStream(file);
		
		XSSFWorkbook wb= new XSSFWorkbook(fis1);
		XSSFSheet sheet= wb.getSheet(sheetnm);
		XSSFRow row= sheet.createRow(rowno);
		XSSFCell cell= row.createCell(colno);
		
		cell.setCellValue(dataval);
		
		FileOutputStream fileout= new FileOutputStream(file);
		wb.write(fileout);
}
}
