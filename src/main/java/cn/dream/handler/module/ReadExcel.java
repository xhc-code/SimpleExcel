package cn.dream.handler.module;

import cn.dream.anno.Excel;
import cn.dream.handler.AbstractExcel;
import cn.dream.handler.WorkbookPropScope;
import cn.dream.handler.bo.SheetData;
import cn.dream.util.ReflectionUtils;
import cn.dream.util.ValueTypeUtils;
import com.sun.xml.internal.ws.addressing.model.ActionNotSupportedException;
import lombok.*;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.Validate;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.beans.BeanUtils;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

public class ReadExcel extends AbstractExcel<ReadExcel> {

    /**
     * 读取指定Sheet名称的Sheet对象，并返回
     * @param sheetName
     * @return
     */
    public ReadExcel readSheet(String sheetName) {
        ReadExcel readExcel = new ReadExcel();
        BeanUtils.copyProperties(this,readExcel, WorkbookPropScope.class);
        readExcel.toggleSheet(sheetName);;
        return readExcel;
    }

    public void toggleSheet(int sheetAt){
        toggleSheet(this.workbook.getSheetName(sheetAt));
    }

    public void toggleSheet(String sheetName){
        Validate.isTrue(this.sheet == null);
        String safeSheetName = validatePassReturnSafeSheetName(sheetName);
        this.sheet = this.workbook.getSheet(safeSheetName);
    }

    @Override
    public ReadExcel newSheet(String sheetName) {
        throw new ActionNotSupportedException("此操作不被支持");
    }

    public void setSheetDataCls(Class<?> dataCls){
        setSheetData(dataCls,new ArrayList<>());
    }

    @Override
    protected void setSheetData(Class<?> dataCls, Collection<?> dataColl) {
        super.setSheetData(dataCls, dataColl);
    }

    public void write(File outputFile) throws IOException {
        throw new ActionNotSupportedException("此操作不被支持");
    }

    public void readData() throws IllegalAccessException {
        SheetData sheetData = getSheetData();

        Excel clsExcel = sheetData.getClsExcel();

        int firstRowNum = this.sheet.getFirstRowNum();
        int lastRowNum = this.sheet.getLastRowNum();

        // 处理Header表头
        int[] headerRowRangeIndex = clsExcel.headerRowRangeIndex();

        // 如果未设置header头的范围,默认是根据有效的第一行并且是有效的第一列的单元格的范围，推算表头的范围
        if(headerRowRangeIndex.length == 0){
            Row row = this.sheet.getRow(firstRowNum);
            CellRangeAddress cellRangeAddress = getCellRangeAddress(this.sheet, row.getCell(row.getFirstCellNum()));
            int firstLastRowIndex = Math.max(0, ObjectUtils.anyNotNull(cellRangeAddress) ? (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow()): 0);

            headerRowRangeIndex = new int[] {firstRowNum,firstRowNum + firstLastRowIndex};
        }

        for(int rowIndex=firstRowNum;rowIndex <= lastRowNum; rowIndex++) {
            if(rowIndex >= headerRowRangeIndex[0] && rowIndex <= headerRowRangeIndex[1]){
                // 处理Header的内容
                putHeaderInfo(this.sheet,headerRowRangeIndex);
                rowIndex += (headerRowRangeIndex[1] - headerRowRangeIndex[0]);
                continue;
            }
            // 处理body data的内容
            putBodyDataByLocation(rowIndex,sheetData.getDataColl());
        }

    }

    /**
     * 将集合转换为指定类型
     * @param ofType
     * @param <T>
     * @return
     */
    public <T> List<T> asType(Class<T> ofType){
        Validate.notNull(ofType);
        SheetData sheetData = getSheetData();
        Class<?> dataCls = sheetData.getDataCls();
        Validate.isTrue(ofType.isAssignableFrom(dataCls),"集合元素的类型不兼容");
        return sheetData.getDataColl().parallelStream().map(v -> (T) v).sequential().collect(Collectors.toList());
    }

    /**
     * 根据列的位置放置数据
     * @param rowIndex
     * @param resultList
     * @throws IllegalAccessException
     */
    public void putBodyDataByLocation(int rowIndex,List<Object> resultList) throws IllegalAccessException {

        SheetData sheetData = getSheetData();

        Class<?> dataCls = sheetData.getDataCls();
        List<Field> fieldList = sheetData.getFieldList();

        Row row = this.sheet.getRow(rowIndex);

        // 按照索引填充数据
        Object newInstance = ReflectionUtils.newInstance(dataCls,false);
        for (int i = 0; i < headerInfoList.size(); i++) {
            HeaderInfo headerInfo = headerInfoList.get(i);
            Field field = fieldList.get(i);

            Cell cell = row.getCell(headerInfo.getColIndex());

            Object cellValue = getMergeCellValue(this.sheet,cell);
            cellValue = ValueTypeUtils.convertValueType(cellValue, field.getType());
            field.set(newInstance,cellValue);

        }
        resultList.add(newInstance);
    }

    private final List<HeaderInfo> headerInfoList = new ArrayList<>();

    /**
     * 往缓存中Put填充Header的基本信息
     * @param sheet
     * @param headerRowRangeIndex
     */
    private void putHeaderInfo(Sheet sheet, int[] headerRowRangeIndex){

        int firstHeaderRowIndex = headerRowRangeIndex[0];
        int lastHeaderRowIndex = headerRowRangeIndex[1];

        Row firstRow = sheet.getRow(firstHeaderRowIndex);

        short firstCellNum = firstRow.getFirstCellNum();
        short lastCellNum = firstRow.getLastCellNum();

        Set<Integer> headerInfoHashCodeSet = new HashSet<>();

        StringBuilder tempStr = new StringBuilder();
        for (short colPointer = firstCellNum; colPointer < lastCellNum; colPointer++) {
            tempStr.setLength(0);
            HeaderInfo.HeaderInfoBuilder headerInfoBuilder = HeaderInfo.builder();

            for (int rowPointer = firstHeaderRowIndex; rowPointer <= lastHeaderRowIndex; rowPointer++) {

                Cell cell = sheet.getRow(rowPointer).getCell(colPointer);
                CellRangeAddress cellRangeAddress = getCellRangeAddress(sheet, cell);
                Object cellValue;
                if(cellRangeAddress != null){
                    Cell firstCell = getFirstCell(this.sheet, cellRangeAddress);
                    cellValue = getCellValue(firstCell);
                    rowPointer+= Math.max(0,(cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow()));

                    headerInfoBuilder.merge(true);
                    headerInfoBuilder.cellAddresses(cellRangeAddress);
                    headerInfoBuilder.cell(firstCell);
                }else {
                    cellValue = getCellValue(cell);

                    headerInfoBuilder.merge(false);
                    headerInfoBuilder.cell(cell);
                }
                if(cellValue != null){
                    tempStr.append(cellValue);
                }
            }

            HeaderInfo headerInfo = headerInfoBuilder.headerName(tempStr.toString())
                    .headerRowRangeIndex(headerRowRangeIndex)
                    .colIndex(colPointer).build();
            if(!headerInfoHashCodeSet.contains(headerInfo.hashCode())){
                headerInfoList.add(headerInfo);
                headerInfoHashCodeSet.add(headerInfo.hashCode());
            }

        }

    }

    @Builder
    @Getter
    @Setter
    @ToString
    @EqualsAndHashCode
    static class HeaderInfo {

        private Object headerName;

        private int[] headerRowRangeIndex;

        @EqualsAndHashCode.Exclude
        private int colIndex;

        private Cell cell;

        private CellRangeAddress cellAddresses;

        private boolean merge;

        public String getHeaderNameAsString(){
            return this.headerName.toString();
        }



    }



    /**
     * 根据Cell获取所处指定Sheet的合并单元格对象
     * @param sheet
     * @param cell
     * @return
     */
    protected CellRangeAddress getCellRangeAddress(Sheet sheet,Cell cell){
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        Optional<CellRangeAddress> cellAddresses = mergedRegions.stream().filter(c -> c.isInRange(cell)).findFirst();
        return cellAddresses.orElse(null);
    }

    private Cell cellRangeAddressToCell(Sheet sheet,CellRangeAddress cellRangeAddress){
        return sheet.getRow(cellRangeAddress.getFirstRow()).getCell(cellRangeAddress.getFirstColumn());
    }


    public static ReadExcel newInstance(Workbook workbook) {
        ReadExcel readExcel = new ReadExcel(workbook);
        return readExcel;
    }

    private ReadExcel(){
        globalCellStyle = workbook.createCellStyle();
    }

    private ReadExcel(Workbook workbook){
        super(true);
        this.workbook = workbook;

        // 初始化操作
        globalCellStyle = workbook.createCellStyle();
    }


}
