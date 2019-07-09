package com.quantumstudio.excelbuilder;


import com.quantumstudio.excelbuilder.ExcelConstants.ExcelFontColor;
import com.quantumstudio.excelbuilder.ExcelConstants.ExcelFontSize;
import com.quantumstudio.excelbuilder.ExcelConstants.ExcelFontType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExcelWorkbook {

	private List<ExcelSheet> excelBook;
	private Workbook workbook;
	private CellStyle currentCellStyle;
	private String fileName;

	public ExcelWorkbook() {
		this.excelBook = new ArrayList<>();
		this.workbook = new XSSFWorkbook();
		withHeaderFont(false, ExcelFontSize.SIZE_11, ExcelFontColor.BLACK, ExcelFontType.ARIAL);
		this.fileName = "prueba" + ".xlsx";
	}

	public ExcelWorkbook addSheet(String name, String[] header, Object[][] contentTable){
		ExcelSheet excelSheet = new ExcelSheet(name, this.currentCellStyle, header, contentTable);
		this.excelBook.add(excelSheet);
		return this;
	}

	public void buildExcel() throws IOException {
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
				row = sheet.createRow(i+1);
				for (int j = 0; j< excelSheet.contentTable[i].length; j++) {
					converterExcelValues(excelSheet.contentTable[i][j], row.createCell(j));
				}
			}
			for(int i = 0; i< excelSheet.header.length; i++){
				sheet.autoSizeColumn(i);
			}
		}

		try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {

			FileOutputStream fos = new FileOutputStream(this.fileName);
			workbook.write(fos);
			fos.close();

			//return outputStream.toByteArray();
		}
	}

	private void converterExcelValues(Object excelValues1, Cell cell1) {
		Cell cell = cell1;

		Object excelValues = excelValues1;
		if(excelValues instanceof String){
			cell.setCellValue(convertToString(excelValues));
		}
		else if(excelValues instanceof Double){
			cell.setCellValue(convertToDouble(excelValues));
		}
		else if(excelValues instanceof Integer){
			cell.setCellValue(convertToInteger(excelValues));
		}
		else if(excelValues instanceof Calendar){
			cell.setCellValue(convertToCalendar(excelValues));
		}
	}

	private String convertToString(Object value) {
		return String.valueOf(value);
	}

	private Double convertToDouble(Object value) {
		return (Double) value;
	}

	private Integer convertToInteger(Object value) {
		return (Integer) value;
	}

	private Calendar convertToCalendar(Object value) {
		return (Calendar) value;
	}

	public ExcelWorkbook withHeaderFont(boolean isBold, ExcelFontSize fontHeight, ExcelFontColor excelFontColor, ExcelFontType fontName){
		Font font = this.workbook.createFont();
		this.currentCellStyle = this.workbook.createCellStyle();
		font.setBold(isBold);
		font.setFontHeightInPoints(fontHeight.getExcelFontSize());
		font.setColor(excelFontColor.getExcelFontsColorIndex());
		font.setFontName(fontName.name());
		this.currentCellStyle.setFont(font);
		return this;
	}

	//TODO: metodo para poner estilos en los datos
	//TODO: metodo con filtros en los headers

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


