package com.elenayalejandro.excelbuilder;

import com.elenayalejandro.excelbuilder.ExcelConstants.ExcelFontColor;
import com.elenayalejandro.excelbuilder.ExcelConstants.ExcelFontSize;
import com.elenayalejandro.excelbuilder.ExcelConstants.ExcelFontType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class ExcelWorkbook {

	private List<ExcelSheet> excelBook;
	private Workbook workbook;
	private CellStyle currentHeaderStyle;
	private CellStyle currentContentFont;
	private boolean isFiltered;
	private String fileName;

	public ExcelWorkbook() {
		this.excelBook = new ArrayList<>();
		this.workbook = new XSSFWorkbook();
		this.isFiltered = false;
		withHeaderFont(true, ExcelFontSize.SIZE_12, ExcelFontColor.BLACK, ExcelFontType.CALIBRI);
		withContentFont(false, ExcelFontSize.SIZE_11, ExcelFontColor.BLACK, ExcelFontType.ARIAL);
		this.fileName = "prueba" + ".xlsx";
	}

	public ExcelWorkbook addSheet(String name, String[] header, Object[][] contentTable){
		ExcelSheet excelSheet = new ExcelSheet(name, this.currentHeaderStyle, this.currentContentFont,  header, contentTable, isFiltered);
		this.excelBook.add(excelSheet);
		return this;
	}

	public OutputStream buildExcel(OutputStream outputStream) throws IOException {
		for(ExcelSheet excelSheet: this.excelBook){
			Sheet sheet = this.workbook.createSheet(excelSheet.name);
			Row row = sheet.createRow(0);

			for (int i = 0; i< excelSheet.header.length; i++){
				Cell cell = row.createCell(i);
				cell.setCellValue(excelSheet.header[i]);
				cell.setCellStyle(excelSheet.userHeaderStyle);
			}

			for (int i = 0; i<excelSheet.contentTable.length; i++) {
				row = sheet.createRow(i+1);
				for (int j = 0; j< excelSheet.contentTable[i].length; j++) {
					Cell cell = row.createCell(j);
					converterExcelValues(excelSheet.contentTable[i][j], row.createCell(j));
					cell.setCellStyle(excelSheet.userContentFont);
				}
			}

			if(excelSheet.isFiltered)
				sheet.setAutoFilter(
						new CellRangeAddress(0, excelSheet.contentTable.length - 1, 0, excelSheet.header.length - 1));

			for(int i = 0; i< excelSheet.header.length; i++){
				sheet.autoSizeColumn(i);
			}
		}
		this.workbook.write(outputStream);
		return outputStream;
	}

	private void converterExcelValues(Object excelValues, Cell cell) {

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
		else if(excelValues instanceof Date){
			cell.setCellValue(convertToDate(excelValues));
		}else {
			throw new IllegalArgumentException("No se admite un tipo que no sea String, Double, Integer, Date o Calendar");
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

	private String convertToCalendar(Object value) {
		DateFormat format = new SimpleDateFormat("dd/MM/yyyy");
		return format.format(((Calendar) value).getTime());
	}

	private String convertToDate(Object value) {
		DateFormat format = new SimpleDateFormat("dd/MM/yyyy");
		return format.format((Date)value);
	}

	public ExcelWorkbook withHeaderFont(boolean isBold, ExcelFontSize fontHeight, ExcelFontColor excelFontColor, ExcelFontType fontName){
		Font font = this.workbook.createFont();
		this.currentHeaderStyle = this.workbook.createCellStyle();
		fontSetters(isBold, fontHeight, excelFontColor, fontName, font);
		this.currentHeaderStyle.setFont(font);
		return this;
	}

	public ExcelWorkbook withContentFont(boolean isBold, ExcelFontSize fontHeight, ExcelFontColor excelFontColor, ExcelFontType fontName) {
		Font font = this.workbook.createFont();
		this.currentContentFont = this.workbook.createCellStyle();
		fontSetters(isBold, fontHeight, excelFontColor, fontName, font);
		this.currentContentFont.setFont(font);
		return this;
	}

	public ExcelWorkbook withHeaderAligment(HorizontalAlignment aligment){
		this.currentHeaderStyle.setAlignment(aligment);
		return this;
	}

	public ExcelWorkbook withContentAligment(HorizontalAlignment aligment){
		this.currentContentFont.setAlignment(aligment);
		return this;
	}

	public ExcelWorkbook addFilter(boolean isFiltered){
		this.isFiltered = isFiltered;
		return this;
	}

	private void fontSetters(boolean isBold, ExcelFontSize fontHeight, ExcelFontColor excelFontColor, ExcelFontType fontName, Font font) {
		font.setBold(isBold);
		font.setFontHeightInPoints(fontHeight.getExcelFontSize());
		font.setColor(excelFontColor.getExcelFontsColorIndex());
		font.setFontName(fontName.name());
	}

	private class ExcelSheet {

		private String name;
		private CellStyle userHeaderStyle;
		private CellStyle userContentFont;
		private String[] header;
		private Object[][] contentTable;
		private boolean isFiltered;

		private ExcelSheet(String name, CellStyle userHeaderStyle, CellStyle userContentFont,  String[] header, Object[][] contentTable, boolean isFiltered) {
			this.isFiltered = isFiltered;
			this.name = name;
			this.userHeaderStyle = userHeaderStyle;
			this.userContentFont = userContentFont;
			this.header = header;
			this.contentTable = contentTable;
		}
	}
}


