package de.xm.jdbcexcel;

import static java.util.Collections.emptyMap;

import de.xm.jdbcexcel.cellwriters.BigDecimalCellWriter;
import de.xm.jdbcexcel.cellwriters.BooleanCellWriter;
import de.xm.jdbcexcel.cellwriters.DateCellWriter;
import de.xm.jdbcexcel.cellwriters.NumberCellWriter;
import de.xm.jdbcexcel.cellwriters.ObjectCellWriter;
import de.xm.jdbcexcel.cellwriters.ReplaceableStringCellWriter;
import de.xm.jdbcexcel.cellwriters.StringCellWriter;
import de.xm.jdbcexcel.tabs.ExcelTab;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.jdbc.core.ArgumentPreparedStatementSetter;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.RowCallbackHandler;
import org.springframework.lang.NonNull;
import org.springframework.util.StringUtils;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.sql.Date;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Types;
import java.util.List;
import java.util.Map;

@Slf4j
public class ExcelWriter {

    public static final int ROWS_IN_MEMORY = 500;
    public static final int MAX_COLUMN_WIDTH_CHARS = 100;

    protected final JdbcTemplate template;
    protected final Map<String, String> stringReplacements;

    public ExcelWriter(JdbcTemplate template) {
        this(template, emptyMap());
    }

    public ExcelWriter(JdbcTemplate template, Map<String, String> replacements) {
        this.template = template;
        this.stringReplacements = replacements;
    }

    public byte[] createExcel(ExcelTab exportTab) throws IOException {
        return createExcel(List.of(exportTab));
    }

    public byte[] createExcel(List<ExcelTab> exportTabs) throws IOException {
        try (SXSSFWorkbook workbook = new SXSSFWorkbook(ROWS_IN_MEMORY)) {
            exportTabs.forEach((tab) -> {
                log.debug("Adding sheet {}", tab.getName());

                SXSSFSheet fieldSheet = workbook.createSheet(tab.getName());

                List<Object> arguments = tab.getParameter();
                String sql = tab.getSql();

                if (checkParameterCount(sql, arguments.size())) {
                    log.debug("Adding {} as parameters to query", arguments);
                    addTab(workbook, fieldSheet, sql, arguments.toArray());
                } else {
                    log.warn("Unable to add sheet {} cause {} are required but {} given", tab, arguments, arguments.size());
                }
            });

            return createByteArray(workbook);
        }
    }

    private boolean checkParameterCount(String sqlStatement, int actualParameterCount) {
        int paramCount = StringUtils.countOccurrencesOf(sqlStatement, "?");
        return paramCount == actualParameterCount;
    }

    private byte[] createByteArray(SXSSFWorkbook workbook) throws IOException {
        try (ByteArrayOutputStream stream = new ByteArrayOutputStream(1_000_000)) {
            workbook.write(stream);

            workbook.dispose();

            return stream.toByteArray();
        }
    }

    private void addTab(SXSSFWorkbook workbook, SXSSFSheet exportSheet, String sql, Object[] arguments) {
        log.info("Adding tab for query '{}' to export", sql);
        template.query(
            sql,
            new ArgumentPreparedStatementSetter(arguments),
            new ExportCallbackHandler(workbook, exportSheet, stringReplacements)
        );
    }

    static class ExportCallbackHandler implements RowCallbackHandler {

        private final SXSSFWorkbook workbook;
        private final SXSSFSheet exportSheet;

        private final DateCellWriter dateCellWriter;
        private final StringCellWriter stringCellWriter;
        private final ReplaceableStringCellWriter replaceableStringCellWriter;
        private final NumberCellWriter numberCellWriter;
        private final BigDecimalCellWriter bigDecimalCellWriter;
        private final ObjectCellWriter objectCellWriter;
        private final BooleanCellWriter booleanCellWriter;

        private int[] maxColumnWidths;

        ExportCallbackHandler(SXSSFWorkbook workbook,
                              SXSSFSheet exportSheet,
                              Map<String, String> replacements) {
            this.workbook = workbook;
            this.exportSheet = exportSheet;

            this.dateCellWriter = new DateCellWriter();
            this.stringCellWriter = new StringCellWriter();
            this.replaceableStringCellWriter = new ReplaceableStringCellWriter(replacements);
            this.numberCellWriter = new NumberCellWriter();
            this.bigDecimalCellWriter = new BigDecimalCellWriter();
            this.objectCellWriter = new ObjectCellWriter();
            this.booleanCellWriter = new BooleanCellWriter();
        }

        @Override
        public void processRow(@NonNull ResultSet rs) throws SQLException {
            ResultSetMetaData metaData = rs.getMetaData();

            int rowIndex = rs.getRow();
            log.trace("Adding row {} to excel sheet", rowIndex);

            if (rowIndex == 1) {
                maxColumnWidths = new int[metaData.getColumnCount()];
                writeHeaderRow(metaData, maxColumnWidths);
            }

            writeRow(rs, metaData, maxColumnWidths);

            if (rs.isLast()) {
                resizeRows(exportSheet, maxColumnWidths);
            }
        }

        private void writeHeaderRow(ResultSetMetaData metaData, int[] maxColumnWidths) throws
            SQLException {

            SXSSFRow headerRow = exportSheet.createRow(0);

            for (int i = 1; i <= metaData.getColumnCount(); i++) {
                String columnHeader = metaData.getColumnLabel(i);

                // header row is always first so we dont have to worry about overwriting something
                maxColumnWidths[i - 1] = columnHeader.length();

                stringCellWriter.writeCellValue(workbook, headerRow, i - 1, columnHeader);
            }
        }

        private void writeRow(ResultSet rs, ResultSetMetaData metaData, int[] maxColumnWidths) throws SQLException {
            SXSSFRow excelRow = exportSheet.createRow(rs.getRow());

            for (int i = 1; i <= metaData.getColumnCount(); i++) {
                maxColumnWidths[i - 1] = Math.max(
                    maxColumnWidths[i - 1],
                    writeCellValueWithWriter(rs, metaData, excelRow, i)
                );
            }
        }

        /**
         * Writes a cell value identified by the current row and columnIndex to the excel sheet
         * and returns the length of the written value in characters
         */
        private int writeCellValueWithWriter(ResultSet rs, ResultSetMetaData metaData, SXSSFRow excelRow, int columnIndex) throws
            SQLException {

            switch (metaData.getColumnType(columnIndex)) {
                case Types.NULL:
                    return 0;

                case Types.VARCHAR:
                case Types.CHAR:
                case Types.LONGVARCHAR:
                case Types.LONGNVARCHAR:
                    String stringCellValue = rs.getString(columnIndex);

                    if (!rs.wasNull()) {
                        return replaceableStringCellWriter.writeCellValue(
                            workbook,
                            excelRow,
                            columnIndex - 1,
                            stringCellValue
                        );
                    }
                    return 0;

                case Types.DATE:
                case Types.TIMESTAMP:
                case Types.TIMESTAMP_WITH_TIMEZONE:
                case Types.TIME:
                case Types.TIME_WITH_TIMEZONE:
                    Date dateCellValue = rs.getDate(columnIndex);

                    if (!rs.wasNull()) {
                        return dateCellWriter.writeCellValue(
                            workbook,
                            excelRow,
                            columnIndex - 1,
                            dateCellValue
                        );
                    }
                    return 0;

                case Types.DOUBLE:
                case Types.INTEGER:
                case Types.SMALLINT:
                case Types.DECIMAL:
                case Types.FLOAT:
                case Types.TINYINT:
                    double doubleCellValue = rs.getDouble(columnIndex);

                    if (!rs.wasNull()) {
                        return numberCellWriter.writeCellValue(
                            workbook,
                            excelRow,
                            columnIndex - 1,
                            doubleCellValue
                        );
                    }
                    return 0;

                case Types.BIGINT:
                case Types.NUMERIC:
                    BigDecimal bigDecimalCellValue = rs.getBigDecimal(columnIndex);

                    if (!rs.wasNull()) {
                        return bigDecimalCellWriter.writeCellValue(
                            workbook,
                            excelRow,
                            columnIndex - 1,
                            bigDecimalCellValue
                        );
                    }
                    return 0;

                case Types.BOOLEAN:
                case Types.BIT:
                    boolean booleanCellValue = rs.getBoolean(columnIndex);

                    if (!rs.wasNull()) {
                        return booleanCellWriter.writeCellValue(
                            workbook,
                            excelRow,
                            columnIndex - 1,
                            booleanCellValue
                        );
                    }
                    return 0;

                case Types.OTHER:
                default:
                    Object objectCellValue = rs.getObject(columnIndex);

                    if (!rs.wasNull()) {
                        return objectCellWriter.writeCellValue(
                            workbook,
                            excelRow,
                            columnIndex - 1,
                            objectCellValue
                        );
                    }
                    return 0;
            }
        }

        private void resizeRows(SXSSFSheet exportSheet, int[] maxColumnWidths) {
            int maxColumnWidth = calculateColumnWidth(MAX_COLUMN_WIDTH_CHARS);

            for (int i = 0; i < maxColumnWidths.length; i++) {
                int columnWidth = Math.min(
                    maxColumnWidth,
                    calculateColumnWidth(maxColumnWidths[i])
                );

                exportSheet.setColumnWidth(i, columnWidth);
            }
        }

        /**
         * Resize columns based on an estimated font width.
         * Since Apache poi uses font measuring (ew), autotracking colun widths gets increasingly
         * expensive if the count of rows grows that the measuring is based on. Additionally, we now have an
         * implicit dependency on fonts being available to base our measurement on.
         * This little formula just estimates the column width, which is good enough for our usecases.
         * <p>
         * https://stackoverflow.com/questions/18983203/how-to-speed-up-autosizing-columns-in-apache-poi#answer-19007294
         */
        private int calculateColumnWidth(int charCount) {
            return ((int) (charCount * 1.14388)) * 256;
        }
    }
}
