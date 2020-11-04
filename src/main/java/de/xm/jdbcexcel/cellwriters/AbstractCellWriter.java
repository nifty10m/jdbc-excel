package de.xm.jdbcexcel.cellwriters;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public abstract class AbstractCellWriter<T> implements CellWriter<T> {
    protected static final DataFormatter DATA_FORMATTER= new DataFormatter();


    @Override
    public int writeCellValue(Workbook workbook, Row row, int cellIndex, T cellValue) {
        if (cellValue == null) {
            return 0;
        }
        Cell cell = row.getCell(cellIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        return doWriteCell(workbook, cell, cellValue);
    }

    protected abstract int doWriteCell(Workbook workbook, Cell cell, T value);

}
