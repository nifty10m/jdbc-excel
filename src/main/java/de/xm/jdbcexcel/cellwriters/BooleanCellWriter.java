package de.xm.jdbcexcel.cellwriters;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

public class BooleanCellWriter extends AbstractCellWriter<Boolean> {

    @Override
    protected int doWriteCell(Workbook workbook, Cell cell, Boolean cellValue) {
        cell.setCellValue(cellValue);
        return cellValue ? 4 : 5;
    }
}
