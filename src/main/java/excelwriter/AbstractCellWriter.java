package excelwriter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public abstract class AbstractCellWriter<T> implements CellWriter<T> {

    @Override
    public void writeCellValue(Workbook workbook, Row row, int cellIndex, T cellValue) {
        if (cellValue == null) {
            return;
        }
        Cell cell = row.getCell(cellIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        doWriteCell(workbook, cell, cellValue);
    }

    protected abstract void doWriteCell(Workbook workbook, Cell cell, T value);

}
