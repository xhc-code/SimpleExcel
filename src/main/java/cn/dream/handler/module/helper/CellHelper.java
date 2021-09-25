package cn.dream.handler.module.helper;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.text.ParseException;

/**
 * 单元格工具，Cell工具类
 * @author xiaohuichao
 * @createdDate 2021/9/24 9:19
 */
@Slf4j
@RequiredArgsConstructor
public class CellHelper {

    /**
     * 此Helper操作的目标Sheet对象
     */
    private final Sheet sheet;

    public void writeCellValue(Cell cell,Object value) throws ParseException {
        SetCellValueHelper.ISetCellValue setValueCell = SetCellValueHelper.getSetValueCell(value.getClass());
        setValueCell.setValue(cell,value,null);
    }

    public void writeCellValue(CellRangeAddress cellAddresses, Object value) {
        writeCellValue(this.sheet,cellAddresses,value);
    }


    public Row createRowIfNotExists(int rowIndex) {
        return createRowIfNotExists(this.sheet,rowIndex);
    }

    public Cell getFirstCell(CellRangeAddress cellAddresses) {
        return getCell(cellAddresses.getFirstRow(), cellAddresses.getFirstColumn());
    }
    public Cell getFirstCell(Cell cell) {
        return getCell(cell.getRowIndex(),cell.getColumnIndex());
    }

    public Cell getCell(int row, int col) {
        return sheet.getRow(row).getCell(col);
    }


    /**
     * 写入合并单元格的值
     * @param cellAddresses 合并单元格的范围
     * @param value 合并单元格的值
     * @return
     */
    public static void writeCellValue(Sheet sheet,CellRangeAddress cellAddresses, Object value) {
        int mergedRegionId = -1;
        try {
            // 将合并单元格中的行和列的单元格对象统统创建出来
            for (int rowIndex = 0; rowIndex <= cellAddresses.getLastRow(); rowIndex++) {
                for (int columnIndex = 0; columnIndex <= cellAddresses.getLastColumn(); columnIndex++) {
                    Row row = createRowIfNotExists(sheet,rowIndex);
                    createCellIfNotExists(row,columnIndex);
                }
            }

            // 首行首列不存在，则进行创建
            Row rowIfNotExists = createRowIfNotExists(sheet,cellAddresses.getFirstRow());
            createCellIfNotExists(rowIfNotExists,cellAddresses.getFirstColumn());

            // 获取合并单元格范围中的首个单元格进行设置值
            Cell firstCell = getFirstCell(sheet,cellAddresses);
            SetCellValueHelper.ISetCellValue setValueCell = SetCellValueHelper.getSetValueCell(value.getClass());
            setValueCell.setValue(firstCell,value,null);

            mergedRegionId = sheet.addMergedRegion(cellAddresses);
        } catch (ParseException e) {
            /**
             * 当出现异常时，移除对应的合并区域
             */
            if(mergedRegionId > -1){
                sheet.removeMergedRegion(mergedRegionId);
            }
            log.info("转换类型异常: {}",e.getMessage());
        }
    }

    public static Cell createCellIfNotExists(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        return cell;
    }

    public static Row createRowIfNotExists(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        return row;
    }

    /**
     * 获取Sheet表里的合并单元格的第一个单元格对象
     *
     * @param cellAddresses
     * @return
     */
    public static Cell getFirstCell(Sheet sheet, CellRangeAddress cellAddresses) {
        return getCell(sheet,cellAddresses.getFirstRow(), cellAddresses.getFirstColumn());
    }
    public static Cell getFirstCell(Sheet sheet, Cell cell) {
        return getCell(sheet,cell.getRowIndex(),cell.getColumnIndex());
    }


    /**
     * 获取指定行指定列的单元格对象
     *
     * @param row 行索引；从0开始
     * @param col 列索引；从0开始
     * @return
     */
    public static Cell getCell(Sheet sheet, int row, int col) {
        return sheet.getRow(row).getCell(col);
    }

}
