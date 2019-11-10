package de.xm.jdbcexcel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public interface CellWriter<T> {

    void writeCellValue(Workbook workbook, Row row, int cellIndex, T cellValue);

}
