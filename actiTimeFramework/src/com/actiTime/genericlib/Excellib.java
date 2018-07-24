package com.actiTime.genericlib;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilterInputStream;
import java.io.FilterOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
public class Excellib 
{
String path=".\\testdata\\testdata2.xlsx";
public String getexceldata(String sname,int rownum,int colnum)throws Throwable
{
	FileInputStream fis=new FileInputStream(path);
	Workbook wb=WorkbookFactory.create(fis);
	Sheet sh=wb.getSheet(sname);
	Row row= sh.getRow(rownum);
	String data=row.getCell(colnum).getStringCellValue();
	wb.close();
	return data;
}

public void setexceldata(String sname,int rownum,int colnum,String data)throws Throwable
{
	FileInputStream fis=new FileInputStream(path);
	Workbook wb=WorkbookFactory.create(fis);
	Sheet sh=wb.getSheet(sname);
	Row row= sh.getRow(rownum);
	Cell cel=row.createCell(colnum);
	cel.setCellValue(data);
	FileOutputStream fos=new FileOutputStream(path);
	wb.write(fos);
	wb.close();
}
}

	 

