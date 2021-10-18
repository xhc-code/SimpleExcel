package cn.dream.handler.module;

import cn.dream.anno.*;
import cn.dream.anno.handler.excelfield.DefaultConverterValueAnnoHandler;
import cn.dream.anno.handler.excelfield.DefaultExcelFieldStyleAnnoHandler;
import cn.dream.anno.handler.excelfield.DefaultSelectValueListAnnoHandler;
import cn.dream.anno.handler.excelfield.DefaultWriteValueAnnoHandler;
import cn.dream.enu.HandlerTypeEnum;
import cn.dream.handler.AbstractExcel;
import cn.dream.handler.bo.CellAddressRange;
import cn.dream.handler.bo.RecordDataValidator;
import cn.dream.handler.bo.SheetData;
import cn.dream.util.ReflectionUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Validate;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Supplier;

@Slf4j
public class WriteExcel extends AbstractExcel<WriteExcel> {

    /**
     * 将根据 @ExcelField 注解的 title 生成Header值，简便使用的方式
     */
    public void generateHeader() {
        Validate.notNull(getSheet());

        SheetData sheetData = getSheetData();
        Excel excelAnno = sheetData.getExcelAnno();

        List<Field> fields = sheetData.getFieldList();

        // 将首行取出，与目标Sheet的首行对比，
        int i = excelAnno.rowIndex();
        int newRowNum = getNewRowNum();
        Row row = createRowIfNotExists(getSheet(),getMaxNum(i, newRowNum, 0));
        AtomicInteger columnIndexAtomic = new AtomicInteger(getMaxNum(excelAnno.columnIndex(), row.getFirstCellNum(), 0));
        for (Field field : fields) {

            writeCellAndNoticeCls(getWorkbook(),null,field,
                    () -> createCellIfNotExists(row, columnIndexAtomic.getAndIncrement()),
                    HandlerTypeEnum.HEADER);

        }

    }

    @Override
    public <T> void setSheetData(Class<T> dataCls, List<T> dataList) {
        super.setSheetData(dataCls, dataList);
    }

    /**
     * 生成主体Body
     */
    public void generateBody() {
        Validate.notNull(getSheet());

        final SheetData sheetData = getSheetData();
        Excel excelAnno = sheetData.getExcelAnno();

        int newRowNum = getNewRowNum();

        List<Field> fieldList = sheetData.getFieldList();

        if(fieldList.size() == 0){
            log.warn("Header集合为空,请确认是否预期的行为？");
            return;
        }

        Collection<?> dataColl = sheetData.getDataList();

        AtomicInteger rowIndex = new AtomicInteger(newRowNum);
        AtomicInteger columnIndex = new AtomicInteger();
        dataColl.forEach(v -> {
            columnIndex.set(0);

            /**
             * 处理单元格的一些操作
             */
            AtomicReference<Row> targetSheetRowIfNotExists = new AtomicReference<>(null);
            for (Field field : fieldList) {
                writeCellAndNoticeCls(getWorkbook(),v,field,() -> {
                    targetSheetRowIfNotExists.set(createRowIfNotExists(getSheet(), rowIndex.get()));
                    return createCellIfNotExists(targetSheetRowIfNotExists.get(), columnIndex.getAndIncrement());
                }, HandlerTypeEnum.BODY);
            }


            Row row = targetSheetRowIfNotExists.get();

            if(row != null){
                /**
                 * 设置行样式单元格信息
                 */
                CellStyle globalCellStyle = getGlobalCellStyle(row.getRowStyle());
                ReflectionUtils.newInstance(excelAnno.handleRowStyle()).setRowStyle(globalCellStyle,v,rowIndex.get());
                globalCellStyle = createCellStyleIfNotExists(globalCellStyle);
                row.setRowStyle(globalCellStyle);

            }
            rowIndex.getAndIncrement();
        });
    }


    /**
     * 自定义处理单元格
     * @param iCustomizeCell
     */
    public void handlerCustomizeCellItem(ICustomizeCell iCustomizeCell) {
        Validate.notNull(getSheet(),"请设置Sheet对象");
        iCustomizeCell.customize(getWorkbook(),getSheet(),
                cellStyle -> this.createCellStyleIfNotExists(getWorkbook(),cellStyle),
                cellHelperSupplier.get()
        );



    }


    /**
     * [写入Excel时会调用]
     * 处理单元格的写入数据和调用针对单元格的一些额外的操作
     */
    protected void writeCellAndNoticeCls(Workbook workbook, Object o, Field field, Supplier<Cell> toCellSupplier, HandlerTypeEnum handlerTypeEnum) {
        Validate.notNull(handlerTypeEnum);
        Validate.notNull(field);
        Validate.notNull(toCellSupplier);

        if(HandlerTypeEnum.HEADER == handlerTypeEnum){

        }else if(HandlerTypeEnum.BODY == handlerTypeEnum){
            Validate.notNull(o);
        }

        ExcelField fieldAnnotation = field.getAnnotation(ExcelField.class);

        // 校验是否包含此字段
        if (fieldAnnotation.apply() && !ignoreFieldApplyList.contains(field.getName())) {

            Cell cell = toCellSupplier.get();

            // 设置自动列宽
            recordAutoColumnMap.putIfAbsent(field, CellAddressRange.builder()
                    .firstCol(cell.getColumnIndex())
                    .lastCol(cell.getColumnIndex())
                    .build());


            // 这里记录下来位置，然后放到write的时候进行设置
            // 设置此Cell可选择的值列表
            if(HandlerTypeEnum.BODY == handlerTypeEnum){

                FieldSelectValueConf selectValueConf = fieldAnnotation.selectValueConf();
                FieldConverterValueConf converterValueConf = fieldAnnotation.converterValueConf();
                // Excel下拉框选项的处理
                String s = converterValueConf.valueExpression();
                if(StringUtils.isNotEmpty(s) || selectValueConf.selectValueListCls() != DefaultSelectValueListAnnoHandler.class ){
                    RecordDataValidator recordDataValidator = recordDataValidatorMap.computeIfAbsent(field, field1 -> {
                        Class<? extends DefaultSelectValueListAnnoHandler> selectValueListCls = selectValueConf.selectValueListCls();
                        DefaultSelectValueListAnnoHandler defaultSelectValueListAnnoHandler = ReflectionUtils.newInstance(selectValueListCls);
                        List<String> parseExpression = defaultSelectValueListAnnoHandler.parseExpression(selectValueConf.selectValues());

                        // 是否需要转换值表达式生成对应的下拉框值
                        if(selectValueConf.buildFromValueExpression()){
                            Class<? extends DefaultConverterValueAnnoHandler> converterValueCls = converterValueConf.valueCls();
                            DefaultConverterValueAnnoHandler defaultConverterValueAnnoHandler = ReflectionUtils.newInstance(converterValueCls);
                            Map<String, String> dictDataMap = defaultConverterValueAnnoHandler.parseExpression(converterValueConf.valueExpression());
                            defaultConverterValueAnnoHandler.fillConverterValue(dictDataMap);
                            parseExpression.addAll(dictDataMap.values());
                        }

                        List<String> selectValueListAnnoHandlerSelectValues = defaultSelectValueListAnnoHandler.getSelectValues(parseExpression);
                        SheetData<?> sheetData = getSheetData();
                        return RecordDataValidator.builder()
                                .selectedItems(selectValueListAnnoHandlerSelectValues.toArray(TYPE_STRINGS))
                                .handlerTypeEnum(handlerTypeEnum)
                                .dataCls(sheetData.getDataCls())
                                .field(field)
                                .o(o)
                                .cellAddressRange(
                                        CellAddressRange.builder()
                                                .firstRow(cell.getRowIndex())
                                                .firstCol(cell.getColumnIndex())
                                                .build()
                                ).build();
                    });

                    CellAddressRange cellAddressRange = recordDataValidator.getCellAddressRange();
                    cellAddressRange.setLastRow(cell.getRowIndex());
                    cellAddressRange.setLastCol(cell.getColumnIndex());
                }

                if(fieldAnnotation.mergeConf().mergeCell()){
                    // 记录合并单元格的范围列表
                    String groupName = getMergeCellGroupName(o, field);
                    if(StringUtils.isNotEmpty(groupName)){
                        Integer fieldIndex = pointerLocationMergeCellMap.getOrDefault(field.getName(),0);
                        String joinGroupName = groupName + STRING_DELIMITER + fieldIndex;

                        if(StringUtils.isNotEmpty(joinGroupName)){
                            CellAddressRange cellAddressRange = recordCellAddressRangeMap.get(joinGroupName);
                            if(cellAddressRange == null){
                                pointerLocationMergeCellMap.put(field.getName(),++fieldIndex);
                                cellAddressRange = CellAddressRange.builder().firstCol(cell.getColumnIndex()).firstRow(cell.getRowIndex()).lastCol(cell.getColumnIndex()).build();
                                recordCellAddressRangeMap.put(groupName + STRING_DELIMITER + fieldIndex, cellAddressRange);
                            }
                            cellAddressRange.setLastRow(cell.getRowIndex());
                        }
                    }
                }


            }

            try {
                field.setAccessible(true);
                AtomicReference<Class<?>> classAtomicReference = new AtomicReference<>(field.getType());
                AtomicReference<Object> valueAtomicReference = new AtomicReference<>(null);
                // 这里判断处理的类型是不是 BODY阶段，否则，o参数是为null，取不到值的，相应的默认值也不进行赋值；原因是 HEADER和FOOTER(未来可能存在)是针对注解本身的值进行操作的
                if(HandlerTypeEnum.BODY == handlerTypeEnum){
                    valueAtomicReference.compareAndSet(null,field.get(o));
                    if(StringUtils.isNotBlank(fieldAnnotation.defaultValue())){
                        valueAtomicReference.compareAndSet(null,fieldAnnotation.defaultValue());
                    }

                    // 当字段有值才需要进行转换
                    if(ObjectUtils.isNotEmpty(valueAtomicReference.get())){
                        // 字典转换值
                        FieldConverterValueConf fieldConverterValueConf = fieldAnnotation.converterValueConf();
                        Class<? extends DefaultConverterValueAnnoHandler> converterValueCls = fieldConverterValueConf.valueCls();
                        DefaultConverterValueAnnoHandler defaultConverterValueAnnoHandler = ReflectionUtils.newInstance(converterValueCls);
                        Map<String, String> dictDataMap = defaultConverterValueAnnoHandler.parseExpression(fieldConverterValueConf.valueExpression());
                        defaultConverterValueAnnoHandler.fillConverterValue(dictDataMap);
                        if(!dictDataMap.isEmpty()){
                            if(fieldConverterValueConf.enableMultiValue()){
                                defaultConverterValueAnnoHandler.multiMapping(dictDataMap,classAtomicReference,valueAtomicReference);
                            }else{
                                defaultConverterValueAnnoHandler.simpleMapping(dictDataMap,classAtomicReference,valueAtomicReference);
                            }
                            classAtomicReference.set(valueAtomicReference.get().getClass());
                        }
                    }


                    Class<? extends DefaultWriteValueAnnoHandler> handlerWriteValue = fieldAnnotation.handlerWriteValue();
                    DefaultWriteValueAnnoHandler writeValueAnnoHandler = ReflectionUtils.newInstance(handlerWriteValue);
                    writeValueAnnoHandler.afterHandler(classAtomicReference, valueAtomicReference);

                }else if(HandlerTypeEnum.HEADER == handlerTypeEnum){
                    classAtomicReference.set(String.class);
                    valueAtomicReference.set(fieldAnnotation.name());
                }

                // 设置样式单元格
                FieldCellStyleConf fieldCellStyleConf = fieldAnnotation.cellStyleConf();
                DefaultExcelFieldStyleAnnoHandler defaultExcelFieldStyleAnnoHandler = ReflectionUtils.newInstance(fieldCellStyleConf.cellStyleCls());
                CellStyle globalCellStyle = getGlobalCellStyle();
                defaultExcelFieldStyleAnnoHandler.cellStyle(globalCellStyle,valueAtomicReference.get(),handlerTypeEnum);
                globalCellStyle = createCellStyleIfNotExists(globalCellStyle);
                cell.setCellStyle(globalCellStyle);

                currentHandlerFieldAnno = fieldAnnotation;
                setCellValue(cell, classAtomicReference.get(), valueAtomicReference.get());
            } catch (IllegalAccessException e) {
                log.error("非法访问 {} 字段,需要排查原因",field.getName());
            }
        }

    }


    private WriteExcel(){}

    private WriteExcel(Workbook workbook){
        super();
        this.workbook = workbook;
    }

    /**
     * 每个单独的对象都需要执行一遍这个操作，以便将缓存的操作信息刷新到WorkBook中
     */
    @Override
    public void flushData() {
        writeData(getSheet());
    }

    @Override
    public WriteExcel newSheet(String sheetName) {
        WriteExcel writeExcel = new WriteExcel();
        writeExcel.embeddedObject = true;
        ReflectionUtils.copyPropertiesByAnno(this,writeExcel);
        writeExcel.createSheet(sheetName);
        writeExcel.initConsumer();
        return writeExcel;
    }

    public static WriteExcel newInstance(Workbook workbook) {
        WriteExcel writeExcel = new WriteExcel(workbook);
        writeExcel.initConsumer();
        return writeExcel;
    }

    /**
     * 创建COpyExcel的对象
     * @param fromWorkbook
     * @return
     */
    public CopyExcel newCopyExcel(Workbook fromWorkbook){
        CopyExcel copyExcel = CopyExcel.newInstance(fromWorkbook, getWorkbook());
        setTransferBeTure(copyExcel);
        return copyExcel;
    }

}
