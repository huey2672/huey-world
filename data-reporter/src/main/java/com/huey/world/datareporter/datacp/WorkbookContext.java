package com.huey.world.datareporter.datacp;

import com.huey.world.datareporter.common.util.PoiHelper;
import com.huey.world.datareporter.datacp.cfg.CfgReader;
import com.huey.world.datareporter.datacp.cfg.spec.Cfg;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;

/**
 * @author huey
 */
@Slf4j
@Getter
public class WorkbookContext {

    private Cfg cfg;

    private Map<String, String> fileMapper = new HashMap<>();

    private Map<String, Workbook> workbookMapper = new HashMap<>();

    public WorkbookContext(String cfgPath) {
        init(cfgPath);
    }

    private void init(String cfgPath) {

        cfg = new CfgReader().readCfg(cfgPath);
        log.info("加载配置完成：{}", cfg);

        cfg.getWorkbooks().forEach(workbook -> {

            fileMapper.put(workbook.getAlias(), workbook.getPath());

            workbookMapper.put(workbook.getAlias(), PoiHelper.createWorkbook(workbook.getPath()));

        });

    }

    public Workbook getWorkbook(String alias) {
        return workbookMapper.get(alias);
    }

    public void exit() {
        fileMapper.forEach((alias, filename) -> {
            PoiHelper.writeWorkbook(getWorkbook(alias), filename);
        });
    }

}