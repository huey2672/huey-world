package com.huey.world.datareporter.datacp.worker;

import com.huey.world.datareporter.common.util.PoiHelper;
import com.huey.world.datareporter.datacp.WorkbookContext;
import com.huey.world.datareporter.datacp.cfg.spec.DateRule;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Calendar;

/**
 * @author huey
 */
@Slf4j
public class DateWorker implements Workable {

    private WorkbookContext context;

    public DateWorker(WorkbookContext context) {
        log.info("实例化 DateWorker");
        this.context = context;
    }

    @Override
    public void execute() {
        context.getCfg().getDateRules().forEach(this::setDateValue);
    }

    private void setDateValue(DateRule dateRule) {

        Workbook workbook = context.getWorkbook(dateRule.getWorkbook());

        dateRule.getCells().forEach(cell -> {

            int rowNum = PoiHelper.getRowNum(cell);
            int colNum = PoiHelper.getColNum(cell);
            Cell poiCell = PoiHelper.getCell(workbook, dateRule.getSheet(), rowNum, colNum);

            if (DateRule.YESTERDAY.equals(dateRule.getDate())) {
                Calendar calendar = Calendar.getInstance();
                calendar.add(Calendar.DATE, -1);
                poiCell.setCellValue(calendar);
                poiCell.setCellStyle(PoiHelper.createDateCellStyle(workbook));
            }

        });

    }


}
