package org.excelautomate;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelReadAnotherMethod {

	public static void main(String[] args) throws IOException {
		excelRead("Sheet1",1, 5);
		excelWrite("Sheet1", 5, 8, "balaji");
	}
	public static String excelRead(String sheetname,int i, int j)
	{
	String value=null;
	try
	{
	File f=new File(System.getProperty("user.dir")+"\\src\\test\\resources\\Sample xlsx.xlsx");
	FileInputStream input=new FileInputStream(f);
	XSSFWorkbook workbook=new XSSFWorkbook(input);
	XSSFSheet sheet=workbook.getSheet(sheetname);

	Row row=sheet.getRow(i);
	Cell cell=row.getCell(j);
	CellType cellType=cell.getCellType();
	if(cellType==CellType.STRING)
	{
	value=cell.getStringCellValue();
	System.out.println(value);
	}
	else if(cellType==CellType.NUMERIC)
	{
	if(DateUtil.isCellDateFormatted(cell))
	{
	Date date=cell.getDateCellValue();
	SimpleDateFormat dt=new SimpleDateFormat("dd/MM/yyyy");
	value=dt.format(date);
	System.out.println(value);
	}
	else
	{
	double num=cell.getNumericCellValue();
	long l=(long)num;
	value=String.valueOf(l);
	
	System.out.println(value);
	}
	}
	}catch (FileNotFoundException e) {
	e.printStackTrace();
	} catch (IOException e) {
	e.printStackTrace();
	}
	return value;
	}
	
	public static void excelWrite(String sheetname,int i, int j, String data) throws IOException
	{
		try
		{
		File f=new File(System.getProperty("user.dir")+"\\src\\test\\resources\\Sample xlsx.xlsx");
		FileInputStream input=new FileInputStream(f);
		XSSFWorkbook workbook=new XSSFWorkbook(input);
		XSSFSheet sheet = workbook.getSheet(sheetname);
		
		XSSFRow row=sheet.getRow(i);
		row.createCell(j).setCellValue(data);
		
		FileOutputStream output=new FileOutputStream(f);
		workbook.write(output);
	}
		catch(FileNotFoundException e)
		{
			e.printStackTrace();
		}catch(IOException e)
		{
			e.printStackTrace();
		}
	
	}

}
