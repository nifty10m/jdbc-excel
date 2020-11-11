package com.newcubator.jdbcexcel.cellwriters;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

@Slf4j
public class ObjectCellWriter extends AbstractCellWriter<Object> {

    @Override
    protected int doWriteCell(Workbook workbook, Cell cell, Object cellValue) {
        String stringValue = cellValue.toString();
        cell.setCellValue(stringValue);

        return stringValue.length();
    }
}
