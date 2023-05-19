package org.excelautomate;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Ignore;
import org.testng.annotations.Test;

public class ExcelReadWrite {
	
	//xls--HSSF
	//xlsx--XSSF
	
	@Ignore
	@Test
	public void read() throws IOException 
	{
		File f=new File("C:\\Users\\User\\eclipse-workspace\\ExcelReadWrite\\src\\test\\resources\\Sample xlsx.xlsx ");
//		System.out.println(f.exists());
		FileInputStream input=new FileInputStream(f);
		XSSFWorkbook workbook=new XSSFWorkbook(input);
		XSSFSheet sheet=workbook.getSheetAt(0);
		
		int totalrows=sheet.getPhysicalNumberOfRows();
		for(int i=0;i<totalrows;i++)
		{
			XSSFRow rows=sheet.getRow(i);
			int totalcells=rows.getPhysicalNumberOfCells();
			for(int j=0;j<totalcells;j++)
			{
				XSSFCell cell=rows.getCell(j);
				cell.getCellType();
				
				if(cell.getCellType()==CellType.NUMERIC)
				{
					double numeric=cell.getNumericCellValue();
					System.out.println(numeric+" ");
				
				}else
				{
					String stringValue= cell.getStringCellValue();
					System.out.println(stringValue);
				}
			}
		System.out.println(" " );
		
		}
		
	workbook.close();
	}
	
	@Test
	public void write() throws IOException
	{
		File f=new File("C:\\Users\\User\\eclipse-workspace\\ExcelReadWrite\\src\\test\\resources\\Sample xlsx.xlsx ");
		FileInputStream input=new FileInputStream(f);
		XSSFWorkbook workbook=new XSSFWorkbook(input);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		
//		XSSFRow row = sheet.getRow(0);
//		row.getCell(0).setCellValue("Serial No");
		
		XSSFRow row=sheet.getRow(0);
		row.createCell(8).setCellValue("Balji");
		
		
		FileOutputStream output= new FileOutputStream(f);
		workbook.write(output);
		workbook.close();
		output.close();
	}
	

}


