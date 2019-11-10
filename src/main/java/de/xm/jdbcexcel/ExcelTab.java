package de.xm.jdbcexcel;

import java.util.List;

public interface ExcelTab {

    String getName();

    List<Object> getParameter();

    String getSql();
}
