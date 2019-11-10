package excelwriter;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.util.StringUtils;

import java.util.Collections;
import java.util.Map;

@Slf4j
public class ReplaceableStringCellWriter extends StringCellWriter {

    private final Map<String, String> replacements;

    public ReplaceableStringCellWriter() {
        this(Collections.emptyMap());
    }

    public ReplaceableStringCellWriter(Map<String, String> replacements) {
        this.replacements = replacements;
    }

    @Override
    protected void doWriteCell(Workbook workbook, Cell cell, String cellValue) {
        super.doWriteCell(workbook, cell, replaceAll(cellValue, replacements));
    }

    public String replaceAll(String input, Map<String, String> replacement) {
        String result = input;
        for (Map.Entry<String, String> entry : replacement.entrySet()) {
            result = StringUtils.replace(result, "{" + entry.getKey() + "}", entry.getValue());
        }
        return result;
    }
}
