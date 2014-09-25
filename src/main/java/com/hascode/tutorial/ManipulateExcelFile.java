package com.hascode.tutorial;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ManipulateExcelFile {

	static double[] customer1 = { 12, 44.2, 32 };
	static double[] customer2 = { 24, 3.5, 11 };
	static double[] customer3 = { 17, 33.25, 42 };

	public static void main(final String[] args) throws IOException {
		FileInputStream file = new FileInputStream(new File(ManipulateExcelFile.class.getClassLoader().getResource("sample.xlsx").getFile()));

		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);
		System.out.println("adding customer statistics to sheet '" + sheet.getSheetName() + "'");

		int startRow = 2;
		int startCol = 1;

		fillCustomerData(customer1, startRow, startCol, sheet);
		fillCustomerData(customer2, startRow, ++startCol, sheet);
		fillCustomerData(customer3, startRow, ++startCol, sheet);

		file.close();

		FileOutputStream outFile = new FileOutputStream(new File("/tmp/updated.xlsx"));
		workbook.write(outFile);
		outFile.close();

	}

	private static void fillCustomerData(final double[] data, int row, final int col, final XSSFSheet sheet) {
		String colName = CellReference.convertNumToColString(col);
		String startCell = colName + (row + 1);
		String stopCell = colName + (row + data.length);
		String sumFormula = String.format("SUM(%s:%s)", startCell, stopCell);
		for (int i = 0; i < data.length; i++) {
			Cell cell = sheet.getRow(row).getCell(col);
			cell.setCellValue(data[i]);
			System.out.println("row: " + row + ", col:" + col + ", cell: " + colName + row + ", data: " + data[i]);
			row++;
		}
		System.out.println("adding sum-formula: " + sumFormula);
		Cell sumCell = sheet.getRow(row).getCell(col);
		sumCell.setCellFormula(sumFormula);
	}
}
