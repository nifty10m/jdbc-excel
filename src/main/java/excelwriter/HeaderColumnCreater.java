package excelwriter;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.util.StringUtils;

import java.sql.ResultSetMetaData;
import java.sql.SQLException;

@Slf4j
public class HeaderColumnCreater {

    public void createColumn(ResultSetMetaData metaData, SXSSFRow headerRow, int column, SXSSFWorkbook workbook) throws SQLException {
        String columnName = metaData.getColumnName(column + 1);
        String columnLabel = metaData.getColumnLabel(column + 1);
        String columnHeader = StringUtils.isEmpty(columnLabel) ? columnName : columnLabel;
        log.debug("Adding column {}", columnHeader);
        StringCellWriter headerCellWriter = new StringCellWriter();
        headerCellWriter.writeCellValue(workbook, headerRow, column, columnHeader);
    }
}
