package cn.dream.handler.module;

import cn.dream.handler.AbstractExcel;
import cn.dream.handler.WorkbookPropScope;
import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import org.apache.commons.lang3.Validate;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.beans.BeanUtils;
import sun.awt.OverrideNativeWindowHandle;

import java.io.File;
import java.io.IOException;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Consumer;
import java.util.function.Predicate;

/**
 * CopyExcel 复制Excel的操作
 */
@SuppressWarnings("DuplicatedCode")
public class CopyExcel extends AbstractExcel<CopyExcel> {

    private Workbook fromWorkbook;
    private Sheet fromSheet;


    private CopyExcel(){
        globalCellStyle = workbook.createCellStyle();
    }

    private CopyExcel(Workbook fromWorkbook,Workbook workbook){
        super(true);
        this.fromWorkbook = fromWorkbook;
        this.workbook = workbook;

        globalCellStyle = workbook.createCellStyle();
    }

    public static CopyExcel newInstance(Workbook fromWorkbook,Workbook workbook){
        return new CopyExcel(fromWorkbook, workbook);
    }

    @Getter
    @Setter
    @ToString
    public static class PointData {

        /**
         * 如果是合并单元格，只记录左上角的第一个单元格位置
         */
        private Cell sourceCell;

        private CellRangeAddress sourceCellAddresses;
        private CellRangeAddress toCellAddresses;

        /**
         * 如果是合并单元格，只记录左上角的第一个单元格位置
         */
        private Cell toCell;

        /**
         * 值，只能是 八大包装类型和Date、Calendar对象
         */
        private Object value;

        /**
         * 将要转换为什么类型,需要自己保证类型是兼容，可以转化
         */
        private Class<?> valueTypeCls;

        /**
         * 是否是合并单元格
         */
        private boolean merge;

        public void setValueWithTypeCls(Object value){
            this.setValue(value);
            this.setValueTypeCls(value.getClass());
        }

    }

    /**
     * copy一行，包含合并单元格的
     *
     * @param modelRowIndex  模板的行索引，从0开始
     * @param modelColIndex  model列索引
     * @param targetRowIndex 目标的行索引，从0开始
     * @param targetColIndex target列索引
     */
    public void copyRow(int modelRowIndex, int modelColIndex, int targetRowIndex, int targetColIndex) {
        Validate.notNull(this.fromWorkbook);
        Validate.notNull(this.fromSheet);
        Validate.isTrue(modelRowIndex > -1, "模型行索引必须大于0");
        Validate.isTrue(targetRowIndex > -1, "目标行索引必须大于0");

        Row row = this.fromSheet.getRow(modelRowIndex);
        short firstCellNum = row.getFirstCellNum();
        short lastCellNum = row.getLastCellNum();
        for (int i = (modelColIndex > firstCellNum) ? modelColIndex : firstCellNum; i < lastCellNum; i++) {
            Cell cell = row.getCell(i);

            if (cell == null) {
                continue;
            }

            copyCell(cell, targetRowIndex - modelRowIndex, targetColIndex - modelColIndex);

        }
    }

    /**
     * 使用过的合并单元格集合,为了防止重复添加相同的合并单元格的问题
     */
    private final Set<CellRangeAddress> useCellRangeAddressSet = new HashSet<>();

    /**
     * Copy一个单元格到指定位置(包含对合并单元格的处理)
     * 需要确保 rowNum 和 colNum 的值大于0等于0
     *
     * @param cell   单元格对象，属于Model的单元格对象
     * @param rowIndex 行索引，从0开始，目标单元格放置的位置
     * @param colIndex 列索引，从0开始，目标单元格放置的位置
     */
    public boolean copyCell(Cell cell, int rowIndex, int colIndex) {

        PointData pointData = new PointData();

        CellRangeAddress rangeRegion = getCellRangeAddress(this.fromSheet,cell);
        if (rangeRegion != null) {
            if (useCellRangeAddressSet.contains(rangeRegion)) {
                return false;
            }

            // 处理合并单元格
            // model相对目标行路径位置
            int firstRowIndex = rangeRegion.getFirstRow() + rowIndex;

            CellRangeAddress newCellRangeAddress = rangeRegion.copy();
            newCellRangeAddress.setFirstRow(firstRowIndex);
            newCellRangeAddress.setLastRow((rangeRegion.getLastRow() - rangeRegion.getFirstRow()) + firstRowIndex);

            // 怎么计算列的相对位置
            int firstCellIndex = rangeRegion.getFirstColumn() + colIndex;
            newCellRangeAddress.setFirstColumn(firstCellIndex);
            newCellRangeAddress.setLastColumn((rangeRegion.getLastColumn() - rangeRegion.getFirstColumn()) + firstCellIndex);

            copyRangeAddressCell(rangeRegion, newCellRangeAddress);

            // 处理合并单元格的值
            Cell sourceFirstCell = getFirstCell(this.fromSheet, rangeRegion);
            Object cellValue = getCellValue(sourceFirstCell);

            Cell toFirstCell = getFirstCell(this.sheet, newCellRangeAddress);


            pointData.setValue(cellValue);
            pointData.setValueTypeCls(cellValue.getClass());
            pointData.setSourceCell(sourceFirstCell);
            pointData.setSourceCellAddresses(rangeRegion);
            pointData.setToCell(toFirstCell);
            pointData.setToCellAddresses(newCellRangeAddress);
            pointData.setMerge(true);

            if(pointDataConsumer != null) {
                pointDataConsumer.accept(pointData);
            }

            setCellValue(toFirstCell, pointData.getValueTypeCls(), pointData.getValue());

            if(iHandlerCellStyle != null) {
                CellStyle globalCellStyle = toFirstCell.getCellStyle();
                iHandlerCellStyle.doHandlerCellStyle(pointData, globalCellStyle);
                toFirstCell.setCellStyle(createCellStyleIfNotExists(globalCellStyle));
            }
            this.useCellRangeAddressSet.add(rangeRegion);
            return true;
        } else {
            // 不是 合并单元格 类型
            Row targetCurrentSheetRow = createRowIfNotExists(this.sheet,rowIndex);

            Cell targetCurrentSheetRowCell = createCellIfNotExists(targetCurrentSheetRow,colIndex);

            targetCurrentSheetRowCell.setCellStyle(createCellStyleIfNotExists(cell.getCellStyle()));

            Object cellValue = getCellValue(cell);

            pointData.setValue(cellValue);
            pointData.setValueTypeCls(cellValue.getClass());
            pointData.setSourceCell(cell);
            pointData.setMerge(false);
            pointData.setToCell(targetCurrentSheetRowCell);

            if(pointDataConsumer != null){
                pointDataConsumer.accept(pointData);
            }

            setCellValue(targetCurrentSheetRowCell, pointData.getValueTypeCls(), pointData.getValue());

            if(iHandlerCellStyle != null){
                CellStyle globalCellStyle = targetCurrentSheetRowCell.getCellStyle();
                iHandlerCellStyle.doHandlerCellStyle(pointData,globalCellStyle);
                targetCurrentSheetRowCell.setCellStyle(createCellStyleIfNotExists(globalCellStyle));
            }

            return true;
        }
    }

    private static final Consumer<PointData> POINT_DATA_EMPTY_CONSUMER = pd -> {};

    protected CellStyle createCellStyleIfNotExists(CellStyle cellStyle){
        return super.createCellStyleIfNotExists(cellStyle);
    }

    @Setter
    private Consumer<PointData> pointDataConsumer;
    @Setter
    private IHandlerCellStyle iHandlerCellStyle;

    @FunctionalInterface
    public interface IHandlerCellStyle {

        void doHandlerCellStyle(PointData pointData,CellStyle cellStyle);

    }

    /**
     * 赋值范围地址单元格
     *
     * @param modelRange
     * @param targetRange
     */
    private void copyRangeAddressCell(CellRangeAddress modelRange, CellRangeAddress targetRange) {
        Map<Cell, Cell> cellMap = new LinkedHashMap<>();

        // 循环创建单元格并设置样式
        for (int modelRowId = modelRange.getFirstRow(), modelColId; modelRowId <= modelRange.getLastRow(); modelRowId++) {

            int relatedRowId = modelRowId - modelRange.getFirstRow();

            Row targetRow = createRowIfNotExists(this.sheet, targetRange.getFirstRow() + relatedRowId);

            modelColId = modelRange.getFirstColumn();
            while (modelColId <= modelRange.getLastColumn()) {

                Cell cell = this.fromSheet.getRow(modelRowId).getCell(modelColId);
                if (cell == null) {
                    modelColId++;
                    continue;
                }

                int relatedColId = modelColId - modelRange.getFirstColumn();
                Cell targetCell = createCellIfNotExists(targetRow, targetRange.getFirstColumn() + relatedColId);

                cellMap.put(cell, targetCell);

                modelColId++;
            }
        }

        this.sheet.addMergedRegion(targetRange);


        /**
         * 这里的单元格样式，应始终使用 {@code Workbook.createCellStyle} 进行创建，默认通过 Cell.getCellStyle()获取到的CellStyle对象是全局默认的，不应该修改，会无法正常显示样式
         */
        cellMap.forEach((modelCell, targetCell) -> targetCell.setCellStyle(createCellStyleIfNotExists(modelCell.getCellStyle())));

    }

    protected void toggleSheet(String sheetName) {
        String safeSheetName = validatePassReturnSafeSheetName(sheetName);
        this.sheet = this.workbook.getSheet(safeSheetName);
    }

    public void toggleFromSheet(String sheetName) {
        String safeSheetName = validatePassReturnSafeSheetName(sheetName);
        toggleFromSheet(this.fromWorkbook.getSheetIndex(safeSheetName));
    }

    public void toggleFromSheet(int sheetIndex){
        this.fromSheet = this.fromWorkbook.getSheetAt(sheetIndex);
    }

    @Override
    public CopyExcel newSheet(String sheetName) {
        CopyExcel copyExcel = new CopyExcel();
        BeanUtils.copyProperties(this,copyExcel, WorkbookPropScope.class);
        copyExcel.createSheet(sheetName);
        return copyExcel;
    }

}
