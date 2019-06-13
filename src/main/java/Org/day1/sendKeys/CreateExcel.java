package Org.day1.sendKeys;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcel {
	public static void main(String[] args) throws IOException {
		File F=new File("D:\\Arun\\Selenium\\sendKeys\\Excel\\Book2.xlsx");
		Workbook W=new XSSFWorkbook();
		Sheet S=W.createSheet("new");
		Row r=S.createRow(2);
		Cell c=r.createCell(1);
		c.setCellValue("Arun");
		FileOutputStream O = new FileOutputStream(F);
		W.write(O);
		System.out.println("Super");
		
	}

}
