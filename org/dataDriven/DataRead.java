package org.dataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataRead {
	public static void data_Read() throws IOException {
		File f = new File("C:\\Users\\91784\\OneDrive\\Desktop\\TN RTO CODES.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheet = wb.getSheetAt(0);
		Row row = sheet.getRow(1);
		Cell cell = row.getCell(2);
		String stringCellValue = cell.getStringCellValue();
		System.out.println(stringCellValue);
	}

	public static void readAll_Data() throws IOException {
		File f = new File("C:\\Users\\91784\\OneDrive\\Desktop\\TN RTO CODES.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook wb1 = new XSSFWorkbook(fis);
		Sheet sheet = wb1.getSheet("Sheet1");
		int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		System.out.println("No of Rows : " + physicalNumberOfRows);

		for (int i = 0; i < physicalNumberOfRows; i++) {
			Row row = sheet.getRow(i);

			int physicalNumberOfCells = row.getPhysicalNumberOfCells();
			//System.out.println("No of Cells :"+physicalNumberOfCells);
			for (int j = 0; j < physicalNumberOfCells; j++) {
				Cell cell = row.getCell(j);
				CellType Type = cell.getCellType();

				if (Type.equals(CellType.STRING)) {
					String stringCellValue = cell.getStringCellValue();
					System.out.println(stringCellValue);
				} else {

					double numericCellValue = cell.getNumericCellValue();
					System.out.println(numericCellValue);
				}

			}

		}

	}

	public static void main(String[] args) throws IOException {
		data_Read();
		readAll_Data();

	}

}
