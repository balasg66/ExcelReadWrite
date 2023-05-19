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

public class Excel1 {
	
//	@Ignore
@Test
public void read() throws IOException
{
File f=new File(System.getProperty("user.dir")+"\\src\\test\\resources\\Sample xlsx.xlsx ");
FileInputStream input=new FileInputStream(f);
XSSFWorkbook workbook=new XSSFWorkbook(input);
XSSFSheet sheet=workbook.getSheet("Sheet1");

int totalrows=sheet.getPhysicalNumberOfRows();

for(int i=1;i<totalrows;i++)
{
	XSSFRow row=sheet.getRow(i);
	int totalcells=row.getPhysicalNumberOfCells();
	for(int j=0;j<totalcells;j++)
	{
	XSSFCell cell=row.getCell(j);
	if(cell.getCellType()==CellType.NUMERIC)
	{
		double num=cell.getNumericCellValue();
		System.out.println(num+" ");
	}else
	{
		String str=cell.getStringCellValue();
		System.out.println(str+" ");
	}
	}
	System.out.println(" ");
}
workbook.close();
}

@Ignore
@Test
public void write() throws IOException
{
File f=new File(System.getProperty("user.dir")+"\\\\src\\\\test\\\\resources\\\\Sample xlsx.xlsx ");
FileInputStream input=new FileInputStream(f);
XSSFWorkbook workbook=new XSSFWorkbook(input);
XSSFSheet sheet=workbook.getSheet("Sheet1");

XSSFRow row=sheet.getRow(2);
//row.getCell(1).setCellValue("NAME");

row.createCell(8).setCellValue("Dharu");

FileOutputStream output=new FileOutputStream(f);
workbook.write(output);
workbook.close();
output.close();
}

}
