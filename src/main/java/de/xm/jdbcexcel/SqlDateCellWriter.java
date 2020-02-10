package de.xm.jdbcexcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

import java.sql.Date;

public class SqlDateCellWriter extends AbstractCellWriter<Date> {

    @Override
    protected void doWriteCell(Workbook workbook, Cell cell, Date value) {
        new DateCellWriter().doWriteCell(workbook, cell, value);
    }
}
