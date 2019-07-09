package com.quantumstudio.excelbuilder;

import com.quantumstudio.excelbuilder.ExcelConstants.ExcelFontColor;
import com.quantumstudio.excelbuilder.ExcelConstants.ExcelFontSize;
import com.quantumstudio.excelbuilder.ExcelConstants.ExcelFontType;

import java.io.IOException;
import java.util.Date;

public class Main {

	public static void main(String[] args) throws IOException {
		Object header[] = {"Columna 1", "columna 2"};
		String header2[] = {"Columna 1ggg", "columnagg 2"};
		Object content[][] = {
				{1, 2},
				{"hola", 4.1},
				{30000, new Date()},
				{3, 4},
		};
		new ExcelWorkbook().addSheet("nuevodocumento", header2, content).withHeaderFont(true, ExcelFontSize.SIZE_12
		, ExcelFontColor.AQUA, ExcelFontType.IMPACT).addSheet("hola", header2, content).buildExcel();
	}
}
