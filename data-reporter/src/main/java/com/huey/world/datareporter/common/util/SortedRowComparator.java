package com.huey.world.datareporter.common.util;

import com.huey.world.datareporter.datacp.cfg.spec.SortRule;
import com.huey.world.datareporter.datacp.cfg.spec.SortedRow;
import org.apache.poi.ss.usermodel.CellType;

import java.util.Comparator;

/**
 * @author huey
 */
public class SortedRowComparator implements Comparator<SortedRow> {

    private String order;

    public SortedRowComparator(String order) {
        this.order = order;
    }

    @Override
    public int compare(SortedRow row1, SortedRow row2) {

        int keyColNum = row1.getKeyColNum();
        CellType cellType = row1.getCellTypes().get(keyColNum);

        switch (cellType) {
            case STRING:
                String stringCellValue1 = (String) row1.getCellValues().get(keyColNum);
                String stringCellValue2 = (String) row2.getCellValues().get(keyColNum);
                return SortRule.DESC.equalsIgnoreCase(order) ? stringCellValue2.compareTo(stringCellValue1) : stringCellValue1.compareTo(stringCellValue2);
            case NUMERIC:
                Double numericCellValue1 = (Double) row1.getCellValues().get(keyColNum);
                Double numericCellValue2 = (Double) row2.getCellValues().get(keyColNum);
                return SortRule.DESC.equalsIgnoreCase(order) ? numericCellValue2.compareTo(numericCellValue1) : numericCellValue1.compareTo(numericCellValue2);
            default:
                break;
        }

        return 0;

    }


}
