package de.xm.jdbcexcel;

import static java.util.Collections.emptyMap;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.jdbc.core.ArgumentPreparedStatementSetter;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.RowCallbackHandler;
import org.springframework.lang.NonNull;
import org.springframework.util.ClassUtils;
import org.springframework.util.StringUtils;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Slf4j
public class ExcelWriter {

    public static final int ROWS_IN_MEMORY = 500;

    protected final JdbcTemplate template;
    protected final LinkedHashMap<Class<?>, CellWriter> writers;

    public ExcelWriter(JdbcTemplate template) {
        this(template, emptyMap());
    }

    public ExcelWriter(JdbcTemplate template, Map<String, String> replacements) {
        this.template = template;
        writers = new LinkedHashMap<>();
        writers.put(Date.class, new DateCellWriter());
        writers.put(String.class, new ReplaceableStringCellWriter(replacements));
        writers.put(BigDecimal.class, new BigDecimalCellWriter());
        writers.put(Number.class, new NumberCellWriter());
    }

    public byte[] createExcel(ExcelTab exportTab) throws IOException {
        return createExcel(List.of(exportTab));
    }

    public byte[] createExcel(List<ExcelTab> exportTabs) throws IOException {
        SXSSFWorkbook workbook = new SXSSFWorkbook(ROWS_IN_MEMORY);
        exportTabs.forEach((tab) -> {
            log.debug("Adding sheet {}", tab.getName());
            SXSSFSheet fieldSheet = workbook.createSheet(tab.getName());
            fieldSheet.trackAllColumnsForAutoSizing();
            List<Object> arguments = tab.getParameter();

            String sqlStatement = tab.getSql();
            int paramCount = 0;
            for (int start = sqlStatement.indexOf('?'); start >= 0; start = sqlStatement.indexOf('?', start + 1)) {
                paramCount++;
            }
            if (paramCount == arguments.size()) {
                log.debug("Adding {} as parameters to query", arguments);
                addTab(workbook, fieldSheet, sqlStatement, arguments.toArray());
            } else {
                log.warn("Unable to add sheet {} cause {} are required but {} given", tab, arguments, paramCount);
            }

        });

        byte[] byteArray = createByteArray(workbook);
        workbook.dispose();
        return byteArray;
    }

    protected byte[] createByteArray(SXSSFWorkbook workbook) throws IOException {
        ByteArrayOutputStream stream = new ByteArrayOutputStream(1_000_000);
        workbook.write(stream);
        return stream.toByteArray();
    }

    protected void addTab(SXSSFWorkbook workbook, SXSSFSheet exportSheet, String sql, Object[] arguments) {
        log.info("Adding tab for query '{}' to export", sql);
        template.query(
            sql,
            new ArgumentPreparedStatementSetter(arguments),
            new ExportCallbackHandler(workbook, exportSheet, writers)
        );
    }

    static class ExportCallbackHandler implements RowCallbackHandler {

        private final SXSSFWorkbook workbook;
        private final SXSSFSheet exportSheet;
        final LinkedHashMap<Class<?>, CellWriter> writers;

        ExportCallbackHandler(SXSSFWorkbook workbook, SXSSFSheet exportSheet, LinkedHashMap<Class<?>, CellWriter> writers) {
            this.workbook = workbook;
            this.exportSheet = exportSheet;
            this.writers = writers;
        }

        @Override
        public void processRow(@NonNull ResultSet rs) throws SQLException {
            ResultSetMetaData metaData = rs.getMetaData();
            int row = rs.getRow();
            log.trace("Adding row {} to excel sheet", row);
            SXSSFRow excelRow = exportSheet.createRow(row);
            int columns = metaData.getColumnCount();
            SXSSFRow headerRow = null;
            if (row == 1) {
                headerRow = exportSheet.createRow(0);
            }
            for (int i = 0; i < columns; i++) {
                writeRow(rs, metaData, row, excelRow, headerRow, i);
            }
            if (row == 499 || (row < 499 && rs.isLast())) {
                log.info("Resizing first {} rows ", row);
                resizeRows(columns, exportSheet);
            }
        }

        private <T> void writeRow(@NonNull ResultSet rs, ResultSetMetaData metaData, int row, SXSSFRow excelRow, SXSSFRow headerRow, int i) throws SQLException {
            if (row == 1) {
                createHeaderColumn(metaData, headerRow, i, workbook);
            }
            String className = metaData.getColumnClassName(i + 1);
            Class<T> clazz = resolveClassByName(className);
            T object = rs.getObject(i + 1, clazz);
            CellWriter<T> writer = findCellWriter(clazz);
            log.trace("Found {} for {}", writer.getClass().getSimpleName(), clazz);
            writer.writeCellValue(workbook, excelRow, i, object);
        }

        @SuppressWarnings("unchecked")
        private <T> Class<T> resolveClassByName(String className) {
            return (Class<T>) ClassUtils.resolveClassName(className, this.getClass().getClassLoader());
        }

        protected void createHeaderColumn(ResultSetMetaData metaData, SXSSFRow headerRow, int column, SXSSFWorkbook workbook) throws SQLException {
            String columnName = metaData.getColumnName(column + 1);
            String columnLabel = metaData.getColumnLabel(column + 1);
            String columnHeader = StringUtils.isEmpty(columnLabel) ? columnName : columnLabel;
            log.debug("Adding column {}", columnHeader);
            StringCellWriter headerCellWriter = new StringCellWriter();
            headerCellWriter.writeCellValue(workbook, headerRow, column, columnHeader);
        }

        protected void resizeRows(int columns, SXSSFSheet exportSheet) {
            for (int i = 0; i < columns; i++) {
                exportSheet.autoSizeColumn(i);
            }
            exportSheet.untrackAllColumnsForAutoSizing();
        }

        @SuppressWarnings("unchecked")
        private <T> CellWriter<T> findCellWriter(Class<T> clazz) {
            return writers.entrySet().stream()
                .filter(entry -> entry.getKey().isAssignableFrom(clazz))
                .map(Map.Entry::getValue)
                .findFirst()
                .orElseGet(ObjectCellWriter::new);
        }

    }

}
