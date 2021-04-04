package com.DataDriven_10Am_Batch;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Write_Data {

	public static void write() throws Throwable {

		File f = new File("C:\\Users\\ELCOT\\Desktop\\Amazon_UserData.xlsx");

		FileInputStream fis = new FileInputStream(f);

		Workbook wb = new XSSFWorkbook(fis);

		Sheet createSheet = wb.createSheet("User_Details");

		Row createRow = createSheet.createRow(0);

		Cell createCell = createRow.createCell(0);

		createCell.setCellValue("Username");

		wb.getSheet("User_Details").getRow(0).createCell(1).setCellValue("Password");

		FileOutputStream fos = new FileOutputStream(f);

		wb.write(fos);

		wb.close();
		
		System.out.println("Data Created");

	}

	public static void main(String[] args) throws Throwable {

		write();
		
	}

}
