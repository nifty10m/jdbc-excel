package com.newcubator.jdbcexcel.cellwriters;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Date;

public class DateCellWriter extends AbstractCellWriter<Date> {

    private final static String DATE_FORMAT = "dd.MM.yyyy";

    private CellStyle dateCellStyle;

    @Override
    public int doWriteCell(Workbook workbook, Cell cell, Date cellValue) {
        cell.setCellValue(cellValue);

        if (dateCellStyle == null) {
            dateCellStyle = createNewDateCellStyle(workbook);
        }

        cell.setCellStyle(dateCellStyle);

        return DATE_FORMAT.length();
    }

    private CellStyle createNewDateCellStyle(Workbook workbook) {
        CellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat(workbook.createDataFormat().getFormat(DATE_FORMAT));

        return dateCellStyle;
    }
}
