package com.huey.world.datareporter.datacp.cfg.spec;

import lombok.Data;

import java.util.List;

/**
 * @author huey
 */
@Data
public class Cfg {

    private List<Workbook> workbooks;

    private List<CpRule> cpRules;

    private List<DateRule> dateRules;

    private List<SortRule> sortRules;

}
