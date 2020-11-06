package de.xm.jdbcexcel.tabs;

import java.util.List;

public interface ExcelTab {

    String getName();

    List<Object> getParameter();

    String getSql();

    static ExcelTab ofSql(String sql) {
        return new DefaultExcelTab("Default", sql, List.of());
    }

    static ExcelTab ofSql(String sql, List<Object> parameters) {
        return new DefaultExcelTab("Default", sql, parameters);
    }

    static ExcelTab of(String name, String sql) {
        return new DefaultExcelTab(name, sql, List.of());
    }

    static ExcelTab of(String name, String sql, List<Object> parameters) {
        return new DefaultExcelTab(name, sql, parameters);
    }
}
