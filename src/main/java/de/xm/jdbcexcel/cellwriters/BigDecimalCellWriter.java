package de.xm.jdbcexcel.cellwriters;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

import java.math.BigDecimal;

public class BigDecimalCellWriter extends AbstractCellWriter<BigDecimal> {

    @Override
    protected int doWriteCell(Workbook workbook, Cell cell, BigDecimal value) {
        cell.setCellValue(value.doubleValue());
        return DATA_FORMATTER.formatCellValue(cell).length();
    }
}
