package com.huey.world.datareporter.common.util;

import com.huey.world.datareporter.common.exception.DataReporterException;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * @author huey
 */
@Slf4j
public class PoiHelper {

    private PoiHelper() {
        super();
    }

    private final static String ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

    /**
     * 创建 Workbook 实例
     *
     * @param filename
     * @return
     */
    public static Workbook createWorkbook(String filename) {

        Workbook workbook = null;

        try (InputStream inputStream = new FileInputStream(new File(filename))) {
            String extension = FilenameUtils.getExtension(filename);

            if ("xls".equals(extension)) {
                workbook = new HSSFWorkbook(inputStream);
            }
            else if ("xlsx".equals(extension)) {
                workbook = new XSSFWorkbook(inputStream);
            }

        }
        catch (IOException e) {
            throw new DataReporterException(e);
        }

        return workbook;

    }

    /**
     * workbook
     *
     * @param workbook
     * @param filename
     */
    public static void writeWorkbook(Workbook workbook, String filename) {

        log.info("更新工作簿 {}", filename);

        try (OutputStream outputStream = new FileOutputStream(new File(filename))) {
            workbook.write(outputStream);
        }
        catch (Exception e) {
            log.error("更新工作簿时异常。", e);
        }

    }

    /**
     * 获取工作表
     *
     * @param workbook
     * @param sheetName
     * @return
     */
    public static Sheet getSheet(Workbook workbook, String sheetName) {

        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new DataReporterException("不存在 [" + sheetName + "] sheet");
        }

        return sheet;

    }

    /**
     * 获取单元行
     *
     * @param sheet
     * @param rowNum
     * @return
     */
    public static Row getRow(Sheet sheet, int rowNum) {
        Row row = sheet.getRow(rowNum);
        if (row == null) {
            row = sheet.createRow(rowNum);
        }
        return row;
    }

    /**
     * 获取单元格
     *
     * @param row
     * @param colNum
     * @return
     */
    public static Cell getCell(Row row, int colNum) {
        Cell cell = row.getCell(colNum);
        if (cell == null) {
            cell = row.createCell(colNum);
        }
        return cell;
    }

    public static Cell getCell(Workbook workbook, String sheetName, int rowNum, int colNum) {
        Sheet sheet = getSheet(workbook, sheetName);
        Row row = getRow(sheet, rowNum);
        Cell cell = getCell(row, colNum);
        return cell;
    }

    /**
     * 获取单元格的行号
     *
     * @param cell
     * @return
     */
    public static int getRowNum(String cell) {
        return Integer.parseInt(cell.substring(1)) - 1;
    }

    /**
     * 获取单元格的列号
     *
     * @param cell
     * @return
     */
    public static int getColNum(String cell) {
        return ALPHABET.indexOf(cell.toUpperCase().charAt(0));
    }

    public static CellStyle createDateCellStyle(Workbook workbook) {

        CellStyle cellStyle = workbook.createCellStyle();
        CreationHelper creationHelper = workbook.getCreationHelper();
        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy/M/d"));

        return cellStyle;

    }

}
