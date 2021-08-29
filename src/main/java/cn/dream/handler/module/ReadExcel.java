package cn.dream.handler.module;

import cn.dream.anno.Excel;
import cn.dream.anno.ExcelField;
import cn.dream.anno.handler.excelfield.DefaultConverterValueAnnoHandler;
import cn.dream.handler.AbstractExcel;
import cn.dream.handler.bo.SheetData;
import cn.dream.util.ReflectionUtils;
import cn.dream.util.ValueTypeUtils;
import com.sun.xml.internal.ws.addressing.model.ActionNotSupportedException;
import lombok.*;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Validate;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.atomic.AtomicReference;
import java.util.stream.Collectors;

public class ReadExcel extends AbstractExcel<ReadExcel> {


    public void toggleSheet(int sheetAt){
        toggleSheet(getWorkbook().getSheetName(sheetAt));
    }

    public void toggleSheet(String sheetName){
        Validate.isTrue(this.sheet == null);
        String safeSheetName = validatePassReturnSafeSheetName(sheetName);
        this.sheet = getWorkbook().getSheet(safeSheetName);
    }

    @Override
    public ReadExcel newSheet(String sheetName) {
        throw new ActionNotSupportedException("此操作不被支持");
    }

    public void setSheetDataCls(Class<?> dataCls){
        setSheetData(dataCls,new ArrayList<>());
    }

    @Override
    protected <T> void setSheetData(Class<T> dataCls, List<T> dataList) {
        super.setSheetData(dataCls, dataList);
    }

    public void write(File outputFile) throws IOException {
        throw new ActionNotSupportedException("此操作不被支持");
    }

    public void readData() throws IllegalAccessException {
        SheetData sheetData = getSheetData();

        Excel clsExcel = sheetData.getExcelAnno();

        int dataFirstRowIndex = clsExcel.dataFirstRowIndex();



        int firstRowNum = getSheet().getFirstRowNum();
        int lastRowNum = getSheet().getLastRowNum();

        // 处理Header表头
        int[] headerRowRangeIndex = clsExcel.headerRowRangeIndex();

        // 如果未设置header头的范围,默认是根据有效的第一行并且是有效的第一列的单元格的范围，推算表头的范围
        if(headerRowRangeIndex.length == 0){
            Row row = getSheet().getRow(firstRowNum);
            CellRangeAddress cellRangeAddress = getCellRangeAddress(getSheet(), row.getCell(row.getFirstCellNum()));
            int firstLastRowIndex = Math.max(0, ObjectUtils.anyNotNull(cellRangeAddress) ? (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow()): 0);

            headerRowRangeIndex = new int[] {firstRowNum,firstRowNum + firstLastRowIndex};
        }

        // 校验数据首行是否在表头的范围
        Validate.isTrue(dataFirstRowIndex == -1 || (dataFirstRowIndex >= headerRowRangeIndex[0] && dataFirstRowIndex <= headerRowRangeIndex[0]),"当前表头行范围包含当前数据行，这可能不是你所期望的情况;数据起始行: %d - 表头范围: %s",headerRowRangeIndex,Arrays.toString(headerRowRangeIndex) );

        for(int rowIndex=firstRowNum;rowIndex <= lastRowNum; rowIndex++) {
            if(rowIndex >= headerRowRangeIndex[0] && rowIndex <= headerRowRangeIndex[1]){
                // 处理Header的内容
                putHeaderInfo(getSheet(),headerRowRangeIndex);
                // 这里rowIndex 到和单元格的末尾行
                rowIndex += (headerRowRangeIndex[1] - headerRowRangeIndex[0]);

                if(dataFirstRowIndex>-1 && dataFirstRowIndex > rowIndex){
                    rowIndex = dataFirstRowIndex;
                }
                continue;
            }
            // 处理body data的内容
            putBodyDataByLocation(rowIndex);
        }

    }

    /**
     * 将集合转换为指定类型
     * @param <T>
     * @return
     */
    public <T> List<T> getResult(){
        SheetData<T> sheetData = getSheetData();
        return sheetData.getDataList();
    }

    /**
     * 根据列的位置放置数据
     * @param rowIndex
     * @throws IllegalAccessException
     */
    public void putBodyDataByLocation(int rowIndex) throws IllegalAccessException {

        SheetData sheetData = getSheetData();

        Class<?> dataCls = sheetData.getDataCls();
        List<Field> fieldList = sheetData.getFieldList();

        Excel excelAnno = sheetData.getExcelAnno();

        boolean byHeaderName = excelAnno.byHeaderName();
        Map<String, Field> fieldMap = null;

        if(byHeaderName){
            fieldMap = fieldList.stream().collect(Collectors.toMap(field -> {
                ExcelField fieldAnnotation = field.getAnnotation(ExcelField.class);
                return Optional.ofNullable(fieldAnnotation.validateHeaderName()).orElse(fieldAnnotation.name());
            }, field -> field));
        }

        Row row = getSheet().getRow(rowIndex);
        if(row == null){
            return;
        }

        // 按照索引填充数据
        Object newInstance = ReflectionUtils.newInstance(dataCls,false);

        Field field;
        ExcelField fieldAnnotation;
        for (int i = 0; i < headerInfoList.size(); i++) {
            HeaderInfo headerInfo = headerInfoList.get(i);
            if(byHeaderName){
                String headerNameAsString = headerInfo.getHeaderNameAsString();
                field = fieldMap.get(headerNameAsString);
                Validate.notNull(field, "没有找到名称为 %s 的字段对象",headerNameAsString);
                fieldAnnotation = field.getAnnotation(ExcelField.class);
            }else{
                Validate.isTrue( i < fieldList.size() , "当前字段集合不存在索引 %d ,请检查实体与Excel之间的映射",i);
                field = fieldList.get(i);
                fieldAnnotation = field.getAnnotation(ExcelField.class);
                if(fieldAnnotation != null && fieldAnnotation.validateHeader()){
                    String headerName = fieldAnnotation.validateHeaderName();
                    if(StringUtils.isEmpty(headerName)){
                        headerName = fieldAnnotation.name();
                    }
                    Validate.isTrue(headerName.equals(headerInfo.getHeaderNameAsString()),"Header表头不一致(AnnoHeader - ExcelHeader)：%s - %s",headerName,headerInfo.getHeaderNameAsString());
                }
            }

            Cell cell = row.getCell(headerInfo.getColIndex());

            Class<?> fieldType = field.getType();
            if(cell == null){
                continue;
            }
            Object cellValue = getMergeCellValue(getSheet(),cell);

            if(ObjectUtils.isEmpty(cellValue)){
                cellValue = null;
            }

            // 当字段有值才需要进行转换
            if(ObjectUtils.isNotEmpty(cellValue)){
                // 字典转换值
                Class<? extends DefaultConverterValueAnnoHandler> converterValueCls = fieldAnnotation.converterValueCls();
                DefaultConverterValueAnnoHandler defaultConverterValueAnnoHandler = ReflectionUtils.newInstance(converterValueCls);
                Map<String, String> dictDataMap = defaultConverterValueAnnoHandler.parseExpression(fieldAnnotation.converterValueExpression());
                defaultConverterValueAnnoHandler.fillConverterValue(dictDataMap);
                if(!dictDataMap.isEmpty()){
                    // 反转Map，用于从 Excel读取值并转换值
                    dictDataMap = defaultConverterValueAnnoHandler.reverse(dictDataMap);
                    AtomicReference<Object> valueAtomicReference=new AtomicReference<>(cellValue);
                    if(fieldAnnotation.enableConverterMultiValue()){
                        defaultConverterValueAnnoHandler.multiMapping(dictDataMap,new AtomicReference<>(fieldType),valueAtomicReference);
                    }else{
                        defaultConverterValueAnnoHandler.simpleMapping(dictDataMap,new AtomicReference<>(fieldType),valueAtomicReference);
                    }
                    cellValue = valueAtomicReference.get();
                }
            }

            if(ObjectUtils.isNotEmpty(cellValue)){
                try {
                    setLocalThreadExcelField(fieldAnnotation);
                    cellValue = ValueTypeUtils.convertValueType(cellValue, fieldType);
                }finally {
                    clearLocalThreadExcelField();
                }
            }

            field.set(newInstance,cellValue);

            /**
             * 写入和读取都需要设置数据的格式，看看 defaultFormatter怎么样集成最为合适吧
             */

        }
        sheetData.getDataList().add(newInstance);
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
                    Cell firstCell = getFirstCell(getSheet(), cellRangeAddress);
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

    /**
     * 读取指定Sheet名称的Sheet对象，并返回
     * @param sheetName
     * @return
     */
    public ReadExcel readSheet(String sheetName) {
        ReadExcel readExcel = new ReadExcel();
        ReflectionUtils.copyPropertiesByAnno(this,readExcel);
        readExcel.initConsumerData();
        readExcel.toggleSheet(sheetName);
        readExcel.embeddedObject = true;
        return readExcel;
    }

    /**
     * 每个单独的对象都需要执行一遍这个操作，以便将缓存的操作信息刷新到WorkBook中
     */
    @Override
    public void flushData() {
        writeData(getSheet());
    }

    public static ReadExcel newInstance(Workbook workbook) {
        ReadExcel readExcel = new ReadExcel(workbook);
        readExcel.oneInit();
        return readExcel;
    }

    private ReadExcel(){}

    private ReadExcel(Workbook workbook){
        super();
        this.workbook = workbook;

        initConsumerData();
    }

}
