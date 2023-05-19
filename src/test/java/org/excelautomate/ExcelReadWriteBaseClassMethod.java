package org.excelautomate;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

public class ExcelReadWriteBaseClassMethod {
	
	
	
	public static String excelRead(String sheetname,int i,int j)
	{
		String value=null;
		try
		{
			File f=new File(System.getProperty("user.dir")+"\\src\\test\\resources\\Sample xls.xls");
			FileInputStream input=new FileInputStream(f);
			HSSFWorkbook workbook=new HSSFWorkbook(input);
			HSSFSheet sheet=workbook.getSheet(sheetname);
			
			HSSFRow row=sheet.getRow(i);
			HSSFCell cell=row.getCell(j);
			CellType celltype=cell.getCellType();
			if(celltype==CellType.STRING)
			{
				value=cell.getStringCellValue();
				System.out.println(value);
			}else if(celltype==CellType.NUMERIC)
			{
				if(DateUtil.isCellDateFormatted(cell))
				{
					Date date=cell.getDateCellValue();
					SimpleDateFormat dt=new SimpleDateFormat("dd/MM/yyyy");
					value = dt.format(date);
					System.out.println(value);
				}else
				{
					double num=cell.getNumericCellValue();
					long l=(long)num;
					value=String.valueOf(l);
				}
			}
			
		}catch(FileNotFoundException e)
		{
			e.printStackTrace();
		}catch(IOException e)
		{
			e.printStackTrace();
		}
		return value;
	}
public static void excelWrite(String sheetname,int i,int j,String data) throws IOException
{
	try
	{
	File f=new File(System.getProperty("user.dir")+"\\src\\test\\resources\\Sample xls.xls");
	FileInputStream input =new FileInputStream(f);
	HSSFWorkbook workbook=new HSSFWorkbook(input);
	HSSFSheet sheet=workbook.getSheet(sheetname);
	
	HSSFRow row = sheet.getRow(i);
	row.createCell(j).setCellValue(data);
	
	
	FileOutputStream output=new FileOutputStream(f);
	workbook.write(output);
	workbook.close();
	output.close();
}catch(FileNotFoundException e)
	{
	e.printStackTrace();
	}catch(IOException e)
	{
		e.printStackTrace();
	}
}
public static void main(String[] args) throws IOException {
	excelRead("Sheet1", 4, 6);
	excelWrite("Sheet1", 6, 8, "balaji");
}
}