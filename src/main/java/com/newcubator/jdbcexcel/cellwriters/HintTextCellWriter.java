package com.newcubator.jdbcexcel.cellwriters;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

public class HintTextCellWriter extends AbstractCellWriter<String> {

    @Override
    protected int doWriteCell(Workbook workbook, Cell cell, String value) {
        CellStyle cellStyle = workbook.createCellStyle();

        Font font = workbook.createFont();
        font.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
        font.setFontHeightInPoints(Short.parseShort("14"));

        cellStyle.setFont(font);

        cell.setCellStyle(cellStyle);
        cell.setCellValue(value);

        return value.length();
    }
}
