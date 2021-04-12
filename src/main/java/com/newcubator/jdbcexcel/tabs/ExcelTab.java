package com.newcubator.jdbcexcel.tabs;

import java.util.List;

public interface ExcelTab {

    String getName();

    List<Object> getParameter();

    String getSql();

    String getHintText();

    static ExcelTab ofSql(String sql) {
        return new DefaultExcelTab("Default", sql, List.of(), null);
    }

    static ExcelTab ofSql(String sql, String hintText) {
        return new DefaultExcelTab("Default", sql, List.of(), hintText);
    }

    static ExcelTab ofSql(String sql, List<Object> parameters) {
        return new DefaultExcelTab("Default", sql, parameters, null);
    }

    static ExcelTab ofSql(String sql, List<Object> parameters, String hintText) {
        return new DefaultExcelTab("Default", sql, parameters, hintText);
    }

    static ExcelTab of(String name, String sql) {
        return new DefaultExcelTab(name, sql, List.of(), null);
    }

    static ExcelTab of(String name, String sql, String hintText) {
        return new DefaultExcelTab(name, sql, List.of(), hintText);
    }

    static ExcelTab of(String name, String sql, List<Object> parameters) {
        return new DefaultExcelTab(name, sql, parameters, null);
    }

    static ExcelTab of(String name, String sql, List<Object> parameters, String hintText) {
        return new DefaultExcelTab(name, sql, parameters, hintText);
    }
}
