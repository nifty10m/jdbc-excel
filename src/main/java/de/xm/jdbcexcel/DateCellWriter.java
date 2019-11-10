package de.xm.jdbcexcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Date;

public class DateCellWriter extends AbstractCellWriter<Date> {

    @Override
    public void doWriteCell(Workbook workbook, Cell cell, Date cellValue) {
        cell.setCellValue(cellValue);
        CellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat(workbook.createDataFormat().getFormat("dd.MM.yyyy"));
        cell.setCellStyle(dateCellStyle);
    }

}
