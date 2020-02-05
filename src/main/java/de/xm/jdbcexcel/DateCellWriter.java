package de.xm.jdbcexcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Date;

public class DateCellWriter extends AbstractCellWriter<Date> {

    private CellStyle dateCellStyle;

    @Override
    public void doWriteCell(Workbook workbook, Cell cell, Date cellValue) {
        cell.setCellValue(cellValue);
        synchronized (this) {
            if (dateCellStyle == null) {
                dateCellStyle = workbook.createCellStyle();
                dateCellStyle.setDataFormat(workbook.createDataFormat().getFormat("dd.MM.yyyy"));
            }
        }
        cell.setCellStyle(dateCellStyle);
    }

}
