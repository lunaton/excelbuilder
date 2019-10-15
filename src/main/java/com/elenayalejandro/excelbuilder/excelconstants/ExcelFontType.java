package com.elenayalejandro.excelbuilder.excelconstants;

public enum ExcelFontType {
	ARIAL("Arial"),
	ARIAL_BLACK("Arial Black"),
	CALIBRI("Calibri"),
	CAMBRIA("Cambria"),
	CENTURY_GOTHIC("Century Gothic"),
	COURIER_NEW("Courier New"),
	FRANKLIN_GOTHIC("Franklin_Gothic"),
	IMPACT("Impact"),
	NEWS_GOTHIC_MT("News gothic MT"),
	TIMES_NEW_ROMAN("Times New Roman");

	private final String excelFontNames;

	ExcelFontType(String excelFontNames) {
		this.excelFontNames = excelFontNames;
	}

	public String getExcelFontNames() {
		return excelFontNames;
	}

}
