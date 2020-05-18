package org.data;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {

	public static void main(String[] args) throws IOException {
		FileInputStream fin=new FileInputStream(System.getProperty("user.dir")+"\\Excel\\New Microsoft Office Excel Worksheet.xlsx");
		Workbook w=new XSSFWorkbook(fin);
		Sheet sheet = w.getSheet("Sheet1");
		int rows = sheet.getPhysicalNumberOfRows();
		for (int i = 0; i < rows; i++) {
			Row row = sheet.getRow(i);
			int cells = row.getPhysicalNumberOfCells();
			
			for (int j = 0; j < cells; j++) {
				Cell cell = row.getCell(j);
				System.out.println(cell);
				
				
			}
			
			System.out.println("done");
		}
	}

}
