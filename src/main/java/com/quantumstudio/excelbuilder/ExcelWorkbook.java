package com.quantumstudio.excelbuilder;


import com.quantumstudio.excelbuilder.ExcelConstants.ExcelFontColor;
import com.quantumstudio.excelbuilder.ExcelConstants.ExcelFontSize;
import com.quantumstudio.excelbuilder.ExcelConstants.ExcelFontType;
import org.apache.poi.hssf.model.Sheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Set;



public class ExcelWorkbook {

	private Set<ExcelSheet>excelBook;
	private Workbook workbook;
	private CellStyle currentCellStyle;

	public ExcelWorkbook() {
		this.workbook = new XSSFWorkbook();
		withFont(false, ExcelFontSize.SIZE_11, ExcelFontColor.BLACK, ExcelFontType.ARIAL);
	}

	public ExcelWorkbook addSheet(String name, String[] header, Object[][] contentTable){
		ExcelSheet excelSheet = new ExcelSheet(name, this.currentCellStyle, header, contentTable);
		this.excelBook.add(excelSheet);
		return this;
	}

	public byte[] buildExcel(){
		Font headerFont = this.workbook.createFont();
		for(ExcelSheet excelSheet: this.excelBook){
			Sheet sheet = this.workbook.createSheet(excelSheet.name);
			Row row = sheet.createRow(0);
			for (int i = 0; i< excelSheet.header.length; i++){
				Cell cell = row.createCell(i);
				cell.setCellValue(excelSheet.header[i]);
				cell.setCellStyle(excelSheet.cellStyle);
			}
			for (int i = 0; i<excelSheet.contentTable.length; i++) {
				for (int j = 0; j< .; j++) {

					Row row = sheet.createRow(rowNum++);

				row.createCell(0).setCellValue(product.getSku());
				row.createCell(1).setCellValue(product.getName());
				row.createCell(2).setCellValue(product.getManufacturerName());
				row.createCell(3).setCellValue(product.getStock());
				row.createCell(4).setCellValue(product.getAvgPrice());
				row.createCell(5).setCellValue(product.getSoldUnits());
				row.createCell(6).setCellValue(product.getTotal());
			}
		}

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


