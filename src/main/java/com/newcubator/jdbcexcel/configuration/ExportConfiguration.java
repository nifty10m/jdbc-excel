package com.newcubator.jdbcexcel.configuration;

import lombok.Data;
import lombok.RequiredArgsConstructor;

@RequiredArgsConstructor
@Data
public class ExportConfiguration {

    private final boolean autogenerateHyperlinks = true;
}
