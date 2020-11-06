package de.xm.jdbcexcel.cellwriters;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public interface CellWriter<T> {

    int writeCellValue(Workbook workbook, Row row, int cellIndex, T cellValue);

}
