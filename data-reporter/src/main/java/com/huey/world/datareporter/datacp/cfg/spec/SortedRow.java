package com.huey.world.datareporter.datacp.cfg.spec;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.CellType;

import java.util.List;

/**
 * @author huey
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class SortedRow {

    private int rowNum;

    private List<Object> cellValues;

    private List<CellType> cellTypes;

    private int keyColNum;

}
