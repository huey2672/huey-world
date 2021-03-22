package com.huey.world.datareporter.datacp.cfg.spec;

import lombok.Data;

/**
 * @author huey
 */
@Data
public class CpRule {

    private String sourceWorkbook;

    private String sourceSheet;

    private String targetWorkbook;

    private String targetSheet;

    private String startCell;

    private String endCell;

}
