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
public class ExcelRW
{

@Ignore
@Test
public void read() throws IOException
{

File f=new File(System.getProperty("user.dir")+"\\src\\test\\resources\\Sample xlsx.xlsx");
FileInputStream input=new FileInputStream(f);
XSSFWorkbook workbook=new XSSFWorkbook(input);
XSSFSheet sheet=workbook.getSheet("Sheet1");

int totalRows=sheet.getPhysicalNumberOfRows();
for(int i=0;i<totalRows;i++)
{
	XSSFRow row=sheet.getRow(i);
	int totalCells=row.getPhysicalNumberOfCells();
	for(int j=0;j<totalCells;j++)
	{
	XSSFCell cell=row.getCell(j);
	if(cell.getCellType()==CellType.NUMERIC)
	{
	   double numbers=cell.getNumericCellValue();
	   System.out.println(numbers+" ");
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
//@Ignore
@Test
public void write() throws IOException
{
File f=new File(System.getProperty("user.dir")+"\\src\\test\\resources\\Sample xlsx.xlsx");
FileInputStream input=new FileInputStream(f);
XSSFWorkbook workbook=new XSSFWorkbook(input);
XSSFSheet sheet=workbook.getSheet("Sheet1");

//XSSFRow row=sheet.getRow(0);
//row.getCell(0).setCellValue("Serial num ");

XSSFRow row=sheet.getRow(3);
row.createCell(8).setCellValue("Rithu");

FileOutputStream output=new FileOutputStream(f);
workbook.write(output);
workbook.close();
output.close();
}
}