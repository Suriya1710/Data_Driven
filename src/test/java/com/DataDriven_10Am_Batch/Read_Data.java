package com.DataDriven_10Am_Batch;

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

public class Read_Data {

	public static void particular_Data() throws IOException {

		File f = new File("C:\\Eclipse\\Demo\\DataDriven_10Am_Batch\\User_Details.xlsx");

		FileInputStream fis = new FileInputStream(f);

		Workbook wb = new XSSFWorkbook(fis); // up casting

		Sheet sheetAt = wb.getSheetAt(0);

		Row row = sheetAt.getRow(4);

		Cell cell = row.getCell(0);

		CellType cellType = cell.getCellType();

		if (cellType.equals(CellType.STRING)) {

			String stringCellValue = cell.getStringCellValue();

			System.out.println(stringCellValue);

		}

		else if (cellType.equals(CellType.NUMERIC)) {

			double numericCellValue = cell.getNumericCellValue();

			// narrowing type casting

			int value = (int) numericCellValue;

			System.out.println(value);

		}

	}

	public static void all_Data() throws Throwable {

		File f = new File
				("C:\\Eclipse\\Demo\\DataDriven_10Am_Batch\\User_Details.xlsx");

		FileInputStream fis = new FileInputStream(f);

		Workbook wb = new XSSFWorkbook(fis);

		Sheet sheetAt = wb.getSheetAt(0);

		int row_Size = sheetAt.getPhysicalNumberOfRows();

		for (int i = 0; i < row_Size; i++) {

			Row row = sheetAt.getRow(i);

			int cell_Size = row.getPhysicalNumberOfCells();

			for (int j = 0; j < cell_Size; j++) {

				Cell cell = row.getCell(j);

				CellType cellType = cell.getCellType();

				if (cellType.equals(CellType.STRING)) {

					String stringCellValue = cell.getStringCellValue();

					System.out.println(stringCellValue);

				}

				else if (cellType.equals(CellType.NUMERIC)) {

					double numericCellValue = cell.getNumericCellValue();

					int value = (int) numericCellValue;

					System.out.println(value);

				}

			}

		}

	}

	public static void main(String[] args) throws Throwable {

		particular_Data();

		System.out.println("*****ALL DATA*****");

		all_Data();

	}

}
