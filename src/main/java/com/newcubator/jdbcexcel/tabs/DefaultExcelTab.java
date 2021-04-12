package com.newcubator.jdbcexcel.tabs;

import lombok.Value;

import java.util.List;

@Value
class DefaultExcelTab implements ExcelTab {

    String name;
    String sql;
    List<Object> parameter;
    String hintText;
}
