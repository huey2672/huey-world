package com.huey.world.datareporter.datacp.cfg.spec;

import lombok.Data;

import java.util.List;

/**
 * @author huey
 */
@Data
public class DateRule {

    public final static String YESTERDAY = "yesterday";

    private String workbook;

    private String sheet;

    private List<String> cells;

    private String date;

}
