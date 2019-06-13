package Org.day1.sendKeys;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
public static void main(String[] args) throws IOException {
	File loc=new File("D:\\Arun\\Selenium\\sendKeys\\Excel\\Book1.xlsx");
	FileInputStream obj=new FileInputStream(loc);
Workbook w = new XSSFWorkbook (obj);
Sheet s=w.getSheet("dd");
for(int i=0;i<s.getPhysicalNumberOfRows();i++) {
	Row r=s.getRow(i);
	for(int j=0;j<r.getPhysicalNumberOfCells();j++) {
		Cell c =r.getCell(j);
		//System.out.println(c);
		int type = c.getCellType();
		//System.out.println(type);
	if(type==1) {
		String st = c.getStringCellValue();
			System.out.println(st);
		
	} if(type==0) {
		//double dd = c.getNumericCellValue();
		//System.out.println(dd);
		if(DateUtil.isCellDateFormatted(c)) {
		Date D = c.getDateCellValue();
		SimpleDateFormat name = new SimpleDateFormat("dd-MMM-yy");
		String n = name.format(D);
		System.out.println(n);
		}else {
		double d = c.getNumericCellValue();
		long l=(long)d;
		String valueOf = String.valueOf(l);
		System.out.println(valueOf);
	}
	}
	
}

	
}
}}

