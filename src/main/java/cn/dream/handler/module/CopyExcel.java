package cn.dream.handler.module;

import cn.dream.handler.AbstractExcel;
import cn.dream.handler.module.helper.CellHelper;
import cn.dream.util.ReflectionUtils;
import cn.dream.util.anno.Feature.RequireCopy;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import org.apache.commons.lang3.Validate;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.IOException;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
import java.util.function.Consumer;

/**
 * CopyExcel 复制Excel的操作
 */
@SuppressWarnings("DuplicatedCode")
public class CopyExcel extends AbstractExcel<CopyExcel> {

    @RequireCopy
    private Workbook fromWorkbook;
    @RequireCopy
    private Sheet fromSheet;

    private CopyExcel(){
    }

    private CopyExcel(Workbook fromWorkbook,Workbook workbook){
        super();
        this.fromWorkbook = fromWorkbook;
        this.workbook = workbook;
    }

    protected Workbook getFromWorkbook(){
        Validate.notNull(this.fromWorkbook,"当前未设置工作簿对象，请设置Workbook对象");
        return this.fromWorkbook;
    }

    protected Sheet getFromSheet(){
        Validate.notNull(this.fromSheet,"当前未设置Sheet对象,请通过相关API进行设置");
        return this.fromSheet;
    }

    @Override
    public CopyExcel newSheet(String sheetName) {
        CopyExcel copyExcel = new CopyExcel();
        copyExcel.embeddedObject = true;
        copyExcel.createSheet(sheetName);
        ReflectionUtils.copyPropertiesByAnno(this,copyExcel);
        copyExcel.initConsumer();
        return copyExcel;
    }

    public static CopyExcel newInstance(Workbook fromWorkbook,Workbook workbook){
        CopyExcel copyExcel = new CopyExcel(fromWorkbook, workbook);
        copyExcel.initConsumer();
        return copyExcel;
    }

    /**
     * 创建COpyExcel的对象
     * @return
     */
    public WriteExcel newWriteExcel(){
        WriteExcel writeExcel = WriteExcel.newInstance(getWorkbook());
        setTransferBeTure(writeExcel);
        return writeExcel;
    }

    /**
     * 数据点的容器
     */
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
        Validate.notNull(getFromWorkbook());
        Validate.notNull(getFromSheet());
        Validate.isTrue(modelRowIndex > -1, "模型行索引必须大于0");
        Validate.isTrue(targetRowIndex > -1, "目标行索引必须大于0");

        Row row = getFromSheet().getRow(modelRowIndex);
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

        CellRangeAddress rangeRegion = getCellRangeAddress(getFromSheet(),cell);
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
            Cell sourceFirstCell = CellHelper.getFirstCell(getFromSheet(), rangeRegion);
            Object cellValue = getCellValue(sourceFirstCell);

            Cell toFirstCell = CellHelper.getFirstCell(getSheet(), newCellRangeAddress);


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
                CellStyle globalCellStyle = getGlobalCellStyle();
                iHandlerCellStyle.doHandlerCellStyle(pointData, globalCellStyle);
                toFirstCell.setCellStyle(createCellStyleIfNotExists(globalCellStyle));
            }
            this.useCellRangeAddressSet.add(rangeRegion);
            return true;
        } else {
            // 不是 合并单元格 类型
            Row targetCurrentSheetRow = createRowIfNotExists(getSheet(),rowIndex);

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
                CellStyle globalCellStyle = getGlobalCellStyle();
                iHandlerCellStyle.doHandlerCellStyle(pointData,globalCellStyle);
                targetCurrentSheetRowCell.setCellStyle(createCellStyleIfNotExists(globalCellStyle));
            }

            return true;
        }
    }

    protected CellStyle createCellStyleIfNotExists(CellStyle cellStyle){
        return super.createCellStyleIfNotExists(cellStyle);
    }

    /**
     * 位置点处理数据，在写入Workbook之前，可将要写入的值在这个对象里设置，仅在Copy单元格的时候会调用
     */
    @Setter
    @RequireCopy
    private Consumer<PointData> pointDataConsumer;

    /**
     * 设置样式
     */
    @Setter
    @RequireCopy
    private IHandlerCellStyle iHandlerCellStyle;

    /**
     * 处理单元格样式
     */
    @FunctionalInterface
    public interface IHandlerCellStyle {

        /**
         * 处理设置单元格样式
         * @param pointData 数据点信息对象
         * @param cellStyle 全局的样式，可以将样式设置到此对象,自动进行缓存和应用到当前Cell单元格
         */
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

            Row targetRow = createRowIfNotExists(getSheet(), targetRange.getFirstRow() + relatedRowId);

            modelColId = modelRange.getFirstColumn();
            while (modelColId <= modelRange.getLastColumn()) {

                Cell cell = getFromSheet().getRow(modelRowId).getCell(modelColId);
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

        getSheet().addMergedRegion(targetRange);


        /**
         * 这里的单元格样式，应始终使用 {@code Workbook.createCellStyle} 进行创建，默认通过 Cell.getCellStyle()获取到的CellStyle对象是全局默认的，不应该修改，会无法正常显示样式
         */
        cellMap.forEach((modelCell, targetCell) -> targetCell.setCellStyle(createCellStyleIfNotExists(modelCell.getCellStyle())));

    }

    protected void toggleSheet(String sheetName) {
        String safeSheetName = validatePassReturnSafeSheetName(sheetName);
        this.sheet = getWorkbook().getSheet(safeSheetName);
    }

    public void toggleFromSheet(String sheetName) {
        String safeSheetName = validatePassReturnSafeSheetName(sheetName);
        toggleFromSheet(getFromWorkbook().getSheetIndex(safeSheetName));
    }

    public void toggleFromSheet(int sheetIndex){
        this.fromSheet = getFromWorkbook().getSheetAt(sheetIndex);
    }


    @Override
    public File write(File outputFile) throws IOException {
        Validate.isTrue(!transfer,"不能调用此方法,请 调用 write() 方法写入数据");
        super.write(outputFile);
        return outputFile;
    }

    /**
     * 每个单独的对象都需要执行一遍这个操作，以便将缓存的操作信息刷新到WorkBook中
     */
    @Override
    public void flushData() {
        writeData(getSheet());
    }

}
