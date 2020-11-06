package de.xm.jdbcexcel.cellwriters;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Workbook;

import java.net.URI;
import java.net.URISyntaxException;

@Slf4j
public class StringCellWriter extends AbstractCellWriter<String> {

    @Override
    protected int doWriteCell(Workbook workbook, Cell cell, String cellValue) {
        if (cellValue.startsWith("http")) {
            markCellAsHyperlink(workbook, cellValue, cell);
        }

        cell.setCellValue(cellValue);

        return cellValue.length();
    }

    private void markCellAsHyperlink(Workbook workbook, String cellValue, Cell cell) {
        try {
            URI uri = new URI(cellValue);
            CreationHelper creationHelper = workbook.getCreationHelper();
            Hyperlink link = creationHelper.createHyperlink(HyperlinkType.URL);
            link.setAddress(uri.normalize().toString());
            cell.setHyperlink(link);
        } catch (URISyntaxException e) {
            log.warn("Found string '{}' in cell but it seems not to be a valid URI, cell is not marked as hyperlink", cellValue, e);
        }
    }
}
