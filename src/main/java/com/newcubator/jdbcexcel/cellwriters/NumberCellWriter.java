package com.newcubator.jdbcexcel.cellwriters;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

public class NumberCellWriter extends AbstractCellWriter<Number> {

    @Override
    protected int doWriteCell(Workbook workbook, Cell cell, Number cellValue) {
        cell.setCellValue(cellValue.doubleValue());
        return DATA_FORMATTER.formatCellValue(cell).length();
    }
}
