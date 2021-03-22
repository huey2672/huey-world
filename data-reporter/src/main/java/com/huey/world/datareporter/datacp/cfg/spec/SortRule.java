package com.huey.world.datareporter.datacp.cfg.spec;

import lombok.Data;

/**
 * @author huey
 */
@Data
public class SortRule {

    public final static String DESC = "desc";

    private String workbook;

    private String sheet;

    private Integer startRow;

    private Integer endRow;

    private Integer colCount;

    private String keyCol;

    private String order;


}
