package de.xm.jdbcexcel;

import lombok.Value;

import java.util.List;

@Value
class DefaultExcelTab implements ExcelTab {

    String name;
    String sql;
    List<Object> parameter;
}
