package de.xm.jdbcexcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.ref.WeakReference;
import java.util.Date;

public class DateCellWriter extends AbstractCellWriter<Date> {

    private WeakReference<Workbook> workbookWeakReference;
    private CellStyle dateCellStyle;

    @Override
    public void doWriteCell(Workbook workbook, Cell cell, Date cellValue) {
        cell.setCellValue(cellValue);
        synchronized (this) {
            if (dateCellStyle == null || workbookWeakReference == null || workbookWeakReference.get() != workbook) {
                clearOldReference();
                dateCellStyle = workbook.createCellStyle();
                dateCellStyle.setDataFormat(workbook.createDataFormat().getFormat("dd.MM.yyyy"));
                workbookWeakReference = new WeakReference<>(workbook);
            }
        }
        cell.setCellStyle(dateCellStyle);
    }

    private void clearOldReference() {
        dateCellStyle = null;
        if (workbookWeakReference != null) {
            workbookWeakReference.clear();
            workbookWeakReference = null;
        }
    }

}
