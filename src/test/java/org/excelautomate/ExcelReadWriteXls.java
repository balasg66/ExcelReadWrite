package org.excelautomate;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.testng.annotations.Ignore;
import org.testng.annotations.Test;

public class ExcelReadWriteXls {
	
	@Ignore
	@Test
	public void read() throws FileNotFoundException, IOException
	{
		File f=new File(System.getProperty("user.dir")+"\\src\\test\\resources\\Sample xls.xls");
//		System.out.println(f.exists());
		FileInputStream input=new FileInputStream(f);
		HSSFWorkbook workbook=new HSSFWorkbook(input);
		HSSFSheet sheet=workbook.getSheet("Sheet1");
		
		int totalrows=sheet.getPhysicalNumberOfRows();
		for(int i=0;i<totalrows;i++)
		{
			HSSFRow row=sheet.getRow(i);
			int totalcells=row.getPhysicalNumberOfCells();
			for(int j=1;j<totalcells;j++)
			{
				HSSFCell cell=row.getCell(j);
				if(cell.getCellType()==CellType.NUMERIC)
				{
					double number=cell.getNumericCellValue();
					System.out.println(number);
				}else
				{
					String stringValue=cell.getStringCellValue();
					System.out.println(stringValue);
				}
				
			}
		System.out.println(" ");
		}
		
		workbook.close();
	}
	
	@Test
	public void write() throws IOException
	{
		File f=new File("C:\\Users\\User\\eclipse-workspace\\ExcelReadWrite\\src\\test\\resources\\Sample xls.xls");
		FileInputStream input=new FileInputStream(f);
		HSSFWorkbook workbook=new HSSFWorkbook(input);
		HSSFSheet sheet=workbook.getSheet("Sheet1");
		
//		HSSFRow row=sheet.getRow(0);
//		row.createCell(0).setCellValue("Serial No");
		
		HSSFRow row=sheet.getRow(0);
		row.getCell(0).setCellValue("Balajiiii");
		
		FileOutputStream output=new FileOutputStream(f);
		workbook.write(output);
		workbook.close();
		output.close();
	}

}

