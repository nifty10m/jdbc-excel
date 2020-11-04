package de.xm.jdbcexcel;

import static java.util.Collections.emptyMap;

import de.xm.jdbcexcel.cellwriters.BigDecimalCellWriter;
import de.xm.jdbcexcel.cellwriters.CellWriter;
import de.xm.jdbcexcel.cellwriters.DateCellWriter;
import de.xm.jdbcexcel.cellwriters.NumberCellWriter;
import de.xm.jdbcexcel.cellwriters.ObjectCellWriter;
import de.xm.jdbcexcel.cellwriters.ReplaceableStringCellWriter;
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
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Types;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Slf4j
public class ExcelWriter {

    public static final int ROWS_IN_MEMORY = 500;

    protected final JdbcTemplate template;
    protected final LinkedHashMap<Class<?>, CellWriter> writers;
    protected final Map<String, Class<?>> resolvedColumnTypes;

    public ExcelWriter(JdbcTemplate template) {
        this(template, emptyMap());
    }

    public ExcelWriter(JdbcTemplate template, Map<String, String> replacements) {
        this.template = template;
        writers = new LinkedHashMap<>();
        writers.put(java.util.Date.class, new DateCellWriter());
        writers.put(String.class, new ReplaceableStringCellWriter(replacements));
        writers.put(Number.class, new NumberCellWriter());
        writers.put(BigDecimal.class, new BigDecimalCellWriter());

        resolvedColumnTypes = new HashMap<>();
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

    protected boolean checkParameterCount(String sqlStatement, int actualParameterCount) {
        int paramCount = StringUtils.countOccurrencesOf(sqlStatement, "?");
        return paramCount == actualParameterCount;
    }

    protected byte[] createByteArray(SXSSFWorkbook workbook) throws IOException {
        try (ByteArrayOutputStream stream = new ByteArrayOutputStream(1_000_000)) {
            workbook.write(stream);
            return stream.toByteArray();
        }
    }

    protected void addTab(SXSSFWorkbook workbook, SXSSFSheet exportSheet, String sql, Object[] arguments) {
        log.info("Adding tab for query '{}' to export", sql);
        template.query(
            sql,
            new ArgumentPreparedStatementSetter(arguments),
            new ExportCallbackHandler(workbook, exportSheet, writers, resolvedColumnTypes)
        );
    }

    static class ExportCallbackHandler implements RowCallbackHandler {

        private final SXSSFWorkbook workbook;
        private final SXSSFSheet exportSheet;
        final LinkedHashMap<Class<?>, CellWriter> writers;
        final Map<String, Class<?>> resolvedColumnTypes;

        int[] maxColumnWidths;

        ExportCallbackHandler(SXSSFWorkbook workbook,
                              SXSSFSheet exportSheet,
                              LinkedHashMap<Class<?>, CellWriter> writers,
                              Map<String, Class<?>> resolvedColumnTypes) {
            this.workbook = workbook;
            this.exportSheet = exportSheet;
            this.writers = writers;
            this.resolvedColumnTypes = resolvedColumnTypes;
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

        protected void writeHeaderRow(ResultSetMetaData metaData, int[] maxColumnWidths) throws
            SQLException {

            SXSSFRow headerRow = exportSheet.createRow(0);

            for (int i = 1; i <= metaData.getColumnCount(); i++) {
                String columnHeader = metaData.getColumnLabel(i);

                // header row is always first so we dont have to worry about overwriting something
                maxColumnWidths[i - 1] = columnHeader.length();

                CellWriter<String> stringCellWriter = (CellWriter<String>) writers.get(String.class);
                stringCellWriter.writeCellValue(workbook, headerRow, i - 1, columnHeader);
            }
        }

        private void writeRow(ResultSet rs, ResultSetMetaData metaData, int[] maxColumnWidths) throws SQLException {
            SXSSFRow excelRow = exportSheet.createRow(rs.getRow());

            for (int i = 1; i <= metaData.getColumnCount(); i++) {
                maxColumnWidths[i - 1] = Math.max(
                    maxColumnWidths[i - 1],
                    writeColumnValue(rs, metaData, excelRow, i)
                );
            }
        }

        /**
         * Writes a cell value identified by the current row and columnIndex to the excel sheet
         * and returns the length of the written value in characters
         */
        private <T> int writeColumnValue(ResultSet rs, ResultSetMetaData metaData, SXSSFRow excelRow, int columnIndex) throws SQLException {
            T columnValue = getAsType(rs, metaData, columnIndex);
            CellWriter<T> cellWriter = (CellWriter<T>) findCellWriter(columnValue.getClass());
            return cellWriter.writeCellValue(workbook, excelRow, columnIndex - 1, columnValue);
        }

        private <T> T getAsType(ResultSet rs, ResultSetMetaData metaData, int columnIndex) throws SQLException {
            switch (metaData.getColumnType(columnIndex)) {
                case Types.VARCHAR:
                case Types.NULL:
                case Types.CHAR:
                    return (T) rs.getString(columnIndex);

                case Types.DATE:
                case Types.TIMESTAMP:
                case Types.TIMESTAMP_WITH_TIMEZONE:
                case Types.TIME:
                case Types.TIME_WITH_TIMEZONE:
                    return (T) rs.getDate(columnIndex);

                case Types.DOUBLE:
                case Types.INTEGER:
                case Types.SMALLINT:
                case Types.DECIMAL:
                case Types.FLOAT:
                    return (T) (Double) rs.getDouble(columnIndex);

                case Types.BIGINT:
                case Types.NUMERIC:
                    return (T) rs.getBigDecimal(columnIndex);

                default:
                    throw new IllegalArgumentException(String.format(
                        "No type defined for column type %s and Jdbc type %s",
                        metaData.getColumnClassName(columnIndex),
                        metaData.getColumnType(columnIndex)
                    ));
            }
        }

        @SuppressWarnings("unchecked")
        private <T> CellWriter<T> findCellWriter(Class<T> clazz) {
            return writers.computeIfAbsent(clazz,
                absentClazz -> writers.entrySet().stream()
                    .filter(entry -> entry.getKey().isAssignableFrom(absentClazz))
                    .map(Map.Entry::getValue)
                    .findFirst()
                    .orElseGet(ObjectCellWriter::new)
            );
        }

        private void resizeRows(SXSSFSheet exportSheet, int[] maxColumnWidths) {
            for (int i = 0; i < maxColumnWidths.length; i++) {
                exportSheet.setColumnWidth(i, calculateColumnWidth(maxColumnWidths[i]));
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
