package com.quantumstudio.excelbuilder;


import com.quantumstudio.excelbuilder.ExcelConstants.ExcelFontColor;
import com.quantumstudio.excelbuilder.ExcelConstants.ExcelFontSize;
import com.quantumstudio.excelbuilder.ExcelConstants.ExcelFontType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Set;



public class ExcelWorkbook {

	private Set<ExcelSheet>excelBook;
	private Workbook workbook;
	private CellStyle currentCellStyle;
	private String fileName;

	public ExcelWorkbook(String fileName) {
		this.workbook = new XSSFWorkbook();
		withFont(false, ExcelFontSize.SIZE_11, ExcelFontColor.BLACK, ExcelFontType.ARIAL);
		this.fileName = fileName + ".xlsx";
	}

	public ExcelWorkbook addSheet(String name, String[] header, Object[][] contentTable){
		ExcelSheet excelSheet = new ExcelSheet(name, this.currentCellStyle, header, contentTable);
		this.excelBook.add(excelSheet);
		return this;
	}

	public byte[] buildExcel() throws IOException {
		Font headerFont = this.workbook.createFont();
		for(ExcelSheet excelSheet: this.excelBook){
			Sheet sheet = this.workbook.createSheet(excelSheet.name);
			Row row = sheet.createRow(0);
			for (int i = 0; i< excelSheet.header.length; i++){
				Cell cell = row.createCell(i);
				cell.setCellValue(excelSheet.header[i]);
				cell.setCellStyle(excelSheet.cellStyle);
				sheet.autoSizeColumn(excelSheet.header.length);
			}

			for (int i = 1; i<excelSheet.contentTable.length; i++) {
				row = sheet.createRow(i);
				for (int j = 0; j< excelSheet.contentTable.length; j++) {
					Cell cell = row.createCell(j);
				}
				sheet.autoSizeColumn(excelSheet.contentTable.length);
			}
		}

		FileOutputStream fileOut = new FileOutputStream(this.fileName);

		workbook.write(fileOut);
		fileOut.close();
		workbook.close();

		final InputStream fileInputStream = new FileInputStream(fileName);
		byte[] finalExcel = IOUtils.toByteArray(fileInputStream);
		fileInputStream.close();
		// Files.delete(Paths.get(fileName));
		return finalExcel;
	}



	private ExcelWorkbook withFont(boolean isBold, ExcelFontSize fontHeight, ExcelFontColor excelFontColor, ExcelFontType fontName){
		Font font = this.workbook.createFont();
		this.currentCellStyle = this.workbook.createCellStyle();
		font.setBold(isBold);
		font.setFontHeightInPoints(fontHeight.getExcelFontSize());
		font.setColor(excelFontColor.getExcelFontsColorIndex());
		font.setFontName(fontName.name());
		this.currentCellStyle.setFont(font);
		return this;
	}

	private class ExcelSheet {

		private String name;
		private CellStyle cellStyle;
		private String[] header;
		private Object[][] contentTable;

		public ExcelSheet(String name, CellStyle cellStyle, String[] header, Object[][] contentTable) {
			this.name = name;
			this.cellStyle = cellStyle;
			this.header = header;
			this.contentTable = contentTable;
		}
	}
}


