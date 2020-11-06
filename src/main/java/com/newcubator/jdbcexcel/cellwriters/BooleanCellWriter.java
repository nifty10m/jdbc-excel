package com.newcubator.jdbcexcel.cellwriters;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

public class BooleanCellWriter extends AbstractCellWriter<Boolean> {

    private static final int TRUE_CHAR_LENGTH = 4;
    private static final int FALSE_CHAR_LENGTH = 5;

    @Override
    protected int doWriteCell(Workbook workbook, Cell cell, Boolean cellValue) {
        cell.setCellValue(cellValue);
        return cellValue ? TRUE_CHAR_LENGTH : FALSE_CHAR_LENGTH;
    }
}
