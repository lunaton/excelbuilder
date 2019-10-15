package com.elenayalejandro.excelbuilder;

import com.elenayalejandro.excelbuilder.excelconstants.ExcelFontColor;
import com.elenayalejandro.excelbuilder.excelconstants.ExcelFontSize;
import com.elenayalejandro.excelbuilder.excelconstants.ExcelFontType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
import java.util.Locale;

public class Main {

	public static void main(String[] args) throws IOException {
		Object[] header = {"Columna 1", "columna 2"};
		String[] header2 = {"Columna 1ggg", "columnagg 2"};
		Object[][] content = {
				{1, 2},
				{"hola", 4.1},
				{30000, new Date()},
				{Calendar.getInstance(Locale.getDefault()), new Date()},
				{3, 4},
		};
		new ExcelWorkbook()
				.addSheet("nuevodocumento", header2, content)
				.withHeaderFont(true, ExcelFontSize.SIZE_12, ExcelFontColor.AQUA, ExcelFontType.IMPACT)
				.withHeaderAligment(HorizontalAlignment.CENTER)
				.addFilter(true)
				.addSheet("hola", header2, content)
				.buildExcel(new FileOutputStream("C:/Users/Alejandro/Desktop/dfghjkl.xlsx"));
	}
}
