package cn.dream.handler.module.helper;

import cn.dream.excep.ValueParseException;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

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

    /**
     * 写入单元格只值
     * @param cell 单元格对象
     * @param value 值
     * @throws ValueParseException
     */
    public void writeCellValue(Cell cell,Object value) throws ValueParseException {
        SetCellValueHelper.ISetCellValue setValueCell = SetCellValueHelper.getSetValueCell(value.getClass());
        setValueCell.setValue(cell,value,null);
    }

    /**
     * 写入合并单元格值
     * @param cellAddresses 合并单元格
     * @param value 值
     */
    public void writeCellValue(CellRangeAddress cellAddresses, Object value) throws ValueParseException {
        writeCellValue(this.sheet,cellAddresses,value);
    }

    public void setCellStyle(CellRangeAddress cellAddresses, CellStyle cellStyle) {
        setCellStyle(this.sheet,cellAddresses,cellStyle);
    }

    /**
     * 如果指定行不存在，则进行创建
     * @param rowIndex 行索引,起始为0
     * @return
     */
    public Row createRowIfNotExists(int rowIndex) {
        return createRowIfNotExists(this.sheet,rowIndex);
    }

    /**
     * 获取合并单元格的起始单元格
     * @param cellAddresses
     * @return
     */
    public Cell getFirstCell(CellRangeAddress cellAddresses) {
        return getCell(cellAddresses.getFirstRow(), cellAddresses.getFirstColumn());
    }


    /**
     * 获取指定行和列的单元格对象
     * @param row 行索引
     * @param col 列索引
     * @return
     */
    public Cell getCell(int row, int col) {
        return sheet.getRow(row).getCell(col);
    }


    /**
     * 写入合并单元格的值
     * @param cellAddresses 合并单元格的范围
     * @param value 合并单元格的值
     * @return
     */
    public static void writeCellValue(Sheet sheet,CellRangeAddress cellAddresses, Object value) throws ValueParseException {
        // 将合并单元格中的行和列的单元格对象统统创建出来
        for (int rowIndex = cellAddresses.getFirstRow(); rowIndex <= cellAddresses.getLastRow(); rowIndex++) {
            Row row = createRowIfNotExists(sheet,rowIndex);
            for (int columnIndex = cellAddresses.getFirstColumn(); columnIndex <= cellAddresses.getLastColumn(); columnIndex++) {
                createCellIfNotExists(row,columnIndex);
            }
        }

        // 获取合并单元格范围中的首个单元格进行设置值
        Cell firstCell = getFirstCell(sheet,cellAddresses);
        SetCellValueHelper.ISetCellValue setValueCell = SetCellValueHelper.getSetValueCell(value.getClass());
        setValueCell.setValue(firstCell,value,null);

        sheet.addMergedRegion(cellAddresses);
    }

    /**
     * 设置合并单元格样式
     * @param cellAddresses 合并单元格对象
     * @param cellStyle 单元格样式
     */
    public static void setCellStyle(Sheet sheet,CellRangeAddress cellAddresses, CellStyle cellStyle){
        for (int rowIndex = cellAddresses.getFirstRow(); rowIndex <= cellAddresses.getLastRow(); rowIndex++) {
            Row row = createRowIfNotExists(sheet,rowIndex);
            for (int columnIndex = cellAddresses.getFirstColumn(); columnIndex <= cellAddresses.getLastColumn(); columnIndex++) {
                Cell cellIfNotExists = createCellIfNotExists(row, columnIndex);
                cellIfNotExists.setCellStyle(cellStyle);
            }
        }
    }

    /**
     * 如果单元格不存在则创建
     * @param row 行对象
     * @param columnIndex 列索引
     * @return
     */
    public static Cell createCellIfNotExists(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        return cell;
    }

    /**
     * 如果行不存在则创建
     * @param sheet Sheet对象
     * @param rowIndex 行索引
     * @return
     */
    public static Row createRowIfNotExists(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        return row;
    }

    /**
     * 获取Sheet表里的合并单元格的第一个单元格对象
     * @param cellAddresses 合并单元格对象
     * @return
     */
    public static Cell getFirstCell(Sheet sheet, CellRangeAddress cellAddresses) {
        return getCell(sheet,cellAddresses.getFirstRow(), cellAddresses.getFirstColumn());
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
