package com.huey.world.datareporter.datacp.worker;

import com.huey.world.datareporter.common.util.PoiHelper;
import com.huey.world.datareporter.common.util.SortedRowComparator;
import com.huey.world.datareporter.datacp.WorkbookContext;
import com.huey.world.datareporter.datacp.cfg.spec.SortRule;
import com.huey.world.datareporter.datacp.cfg.spec.SortedRow;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;

/**
 * @author huey
 */
@Slf4j
public class SortWorker implements Workable {

    private WorkbookContext context;

    public SortWorker(WorkbookContext context) {
        log.info("实例化 SortWorker");
        this.context = context;
    }

    @Override
    public void execute() {

        context.getCfg().getSortRules().forEach(this::sort);

    }

    private void sort(SortRule sortRule) {

        List<SortedRow> sortedRows = new ArrayList<>();

        Workbook workbook = context.getWorkbook(sortRule.getWorkbook());
        int startRow = sortRule.getStartRow() - 1;
        int endRow = sortRule.getEndRow() - 1;

        for (int row = startRow; row <= endRow; row++) {

            List<Object> cellValues = new ArrayList<>();

            List<CellType> cellTypes = new ArrayList<>();

            for (int col = 0; col < sortRule.getColCount(); col++) {
                Cell cell = PoiHelper.getCell(workbook, sortRule.getSheet(), row, col);
                Object cellValue = null;
                switch (cell.getCellType()) {
                    case STRING:
                        cellValue = cell.getStringCellValue();
                        break;
                    case NUMERIC:
                        cellValue = cell.getNumericCellValue();
                        break;
                    default:
                        break;
                }
                cellValues.add(cellValue);
                cellTypes.add(cell.getCellType());
            }
            SortedRow sortedRow = new SortedRow(row, cellValues, cellTypes, PoiHelper.getColNum(sortRule.getKeyCol()));
            sortedRows.add(sortedRow);
        }

        log.info("排序前：{}", sortedRows);
        sortedRows.sort(new SortedRowComparator(sortRule.getOrder()));
        log.info("排序后：{}", sortedRows);

        for (int row = startRow; row <= endRow; row++) {

            SortedRow sortedRow = sortedRows.get(row);

            for (int col = 0; col < sortRule.getColCount(); col++) {

                Cell cell = PoiHelper.getCell(workbook, sortRule.getSheet(), row, col);

                switch (cell.getCellType()) {
                    case STRING:
                        cell.setCellValue((String) sortedRow.getCellValues().get(col));
                        break;
                    case NUMERIC:
                        cell.setCellValue((Double) sortedRow.getCellValues().get(col));
                        break;
                    default:
                        break;
                }
            }
        }

    }

}
