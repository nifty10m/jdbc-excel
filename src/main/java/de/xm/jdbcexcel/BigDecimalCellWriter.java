package de.xm.jdbcexcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

import java.math.BigDecimal;

public class BigDecimalCellWriter extends AbstractCellWriter<BigDecimal> {

    @Override
    protected void doWriteCell(Workbook workbook, Cell cell, BigDecimal value) {
        cell.setCellValue(value.doubleValue());
    }
}
