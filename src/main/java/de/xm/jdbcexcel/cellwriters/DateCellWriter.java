package de.xm.jdbcexcel.cellwriters;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.ref.WeakReference;
import java.util.Date;

public class DateCellWriter extends AbstractCellWriter<Date> {

    private final static String DATE_FORMAT = "dd.MM.yyyy";

    private WeakReference<Workbook> workbookWeakReference;
    private CellStyle dateCellStyle;

    @Override
    public int doWriteCell(Workbook workbook, Cell cell, Date cellValue) {
        cell.setCellValue(cellValue);

        synchronized (this) {
            if (dateCellStyle == null || workbookWeakReference == null || workbookWeakReference.get() != workbook) {
                clearOldReference();
                dateCellStyle = createNewDateCellStyle(workbook);

                workbookWeakReference = new WeakReference<>(workbook);
            }
        }

        cell.setCellStyle(dateCellStyle);

        return DATE_FORMAT.length();
    }

    private void clearOldReference() {
        dateCellStyle = null;
        if (workbookWeakReference != null) {
            workbookWeakReference.clear();
            workbookWeakReference = null;
        }
    }

    private CellStyle createNewDateCellStyle(Workbook workbook) {
        CellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat(workbook.createDataFormat().getFormat(DATE_FORMAT));

        return dateCellStyle;
    }
}
