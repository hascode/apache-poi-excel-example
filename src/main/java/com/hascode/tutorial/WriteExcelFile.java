package com.hascode.tutorial;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFBorderFormatting;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class WriteExcelFile {

	public static void main(final String[] args) throws Exception {
		SXSSFWorkbook wb = new SXSSFWorkbook();
		Font headerFont = wb.createFont();
		headerFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		CellStyle headerStyle = wb.createCellStyle();
		headerStyle.setFont(headerFont);
		headerStyle.setBorderBottom(HSSFBorderFormatting.BORDER_THIN);
		Sheet sh = wb.createSheet("First sheet");
		for (int rownum = 0; rownum < 20; rownum++) {
			Row row = sh.createRow(rownum);
			String prefix = "";
			if (rownum == 0) {
				prefix = "Header ";
				row.setRowStyle(headerStyle);
			}
			for (int cellnum = 0; cellnum < 10; cellnum++) {
				Cell cell = row.createCell(cellnum);
				String cellValue = new CellReference(cell).formatAsString();
				cell.setCellValue(prefix + cellValue);
			}

		}

		FileOutputStream out = new FileOutputStream("/tmp/myexcel.xlsx");
		wb.write(out);
		out.close();
		wb.dispose();
	}
}
