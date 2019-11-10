package de.xm.jdbcexcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

public class NumberCellWriter extends AbstractCellWriter<Number> {

    @Override
    protected void doWriteCell(Workbook workbook, Cell cell, Number cellValue) {
        cell.setCellValue(cellValue.doubleValue());
    }
}
