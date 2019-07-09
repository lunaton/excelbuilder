package com.quantumstudio.excelbuilder.ExcelConstants;

public enum ExcelFontSize{
	SIZE_8(8),
	SIZE_9(9),
	SIZE_10(10),
	SIZE_11(11),
	SIZE_12(12),
	SIZE_14(14),
	SIZE_16(16),
	SIZE_18(18),
	SIZE_20(20),
	SIZE_22(22),
	SIZE_24(24),
	SIZE_26(26),
	SIZE_28(28),
	SIZE_30(30),
	SIZE_36(36),
	SIZE_48(48),
	SIZE_72(72);

	public final short excelFontSize;

	ExcelFontSize (int excelFontSize){
		this.excelFontSize = (short) excelFontSize;
	}

	public short getExcelFontSize() {
		return excelFontSize;
	}
}
