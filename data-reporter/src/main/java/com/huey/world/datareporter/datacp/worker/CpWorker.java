package com.huey.world.datareporter.datacp.worker;

import com.huey.world.datareporter.common.util.PoiHelper;
import com.huey.world.datareporter.datacp.WorkbookContext;
import com.huey.world.datareporter.datacp.cfg.spec.CpRule;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author huey
 */
@Slf4j
public class CpWorker implements Workable {

    private WorkbookContext context;

    public CpWorker(WorkbookContext context) {
        log.info("实例化 CpWorker");
        this.context = context;
    }

    @Override
    public void execute() {

        log.info("执行 CpWorker");

        context.getCfg().getCpRules().forEach(this::copy);

    }

    private void copy(CpRule cpRule) {

        Workbook sourceWorkbook = context.getWorkbook(cpRule.getSourceWorkbook());
        Workbook targetWorkbook = context.getWorkbook(cpRule.getTargetWorkbook());

        String startCell = cpRule.getStartCell();
        String endCell = cpRule.getEndCell();

        int startRow = PoiHelper.getRowNum(startCell);
        int startCol = PoiHelper.getColNum(startCell);
        int endRow = PoiHelper.getRowNum(endCell);
        int endCol = PoiHelper.getColNum(endCell);

        log.info("开始从 [{}, {}] 复制数据到 [{}, {}]，数据范围：[{} - {}]。",
                cpRule.getSourceWorkbook(), cpRule.getSourceSheet(),
                cpRule.getTargetWorkbook(), cpRule.getTargetSheet(),
                cpRule.getStartCell(), cpRule.getEndCell()
        );

        for (int row = startRow; row <= endRow; row++) {
            for (int col = startCol; col <= endCol; col++) {
                Cell sourceCell = PoiHelper.getCell(sourceWorkbook, cpRule.getSourceSheet(), row, col);
                Cell targetCell = PoiHelper.getCell(targetWorkbook, cpRule.getTargetSheet(), row, col);
                Object cellValue = copyCell(sourceCell, targetCell);
                log.info("开始复制数据，Cell[{}, {}] = {}。", row, col, cellValue);
            }
        }

    }

    private Object copyCell(Cell sourceCell, Cell targetCell) {

        Object cellValue = null;

        switch (sourceCell.getCellType()) {
            case STRING:
                cellValue = sourceCell.getStringCellValue();
                targetCell.setCellValue((String) cellValue);
                break;
            case NUMERIC:
                cellValue = sourceCell.getNumericCellValue();
                targetCell.setCellValue((Double) cellValue);
                break;
            default:
                break;
        }

        return cellValue;

    }

}
