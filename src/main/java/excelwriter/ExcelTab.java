package excelwriter;

import java.util.List;

public interface ExcelTab {

    String getName();

    List<Object> getParameter();

    String getSql();
}
