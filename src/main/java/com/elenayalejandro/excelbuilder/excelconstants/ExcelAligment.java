package com.elenayalejandro.excelbuilder.excelconstants;

public enum ExcelAligment {

	ALIGN_GENERAL(0),
	ALIGN_LEFT(1),
	ALIGN_CENTER(2),
	ALIGN_RIGHT(3),
	ALIGN_FILL(4),
	ALIGN_JUSTIFY(5),
	ALIGN_CENTER_SELECTION(6);

	public final int aligment;

	ExcelAligment(int aligment) {
		this.aligment = aligment;
	}
}
