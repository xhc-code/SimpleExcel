package cn.dream.handler;

import cn.dream.anno.ExcelField;
import cn.dream.anno.MergeField;
import cn.dream.anno.handler.excelfield.*;
import cn.dream.enu.HandlerTypeEnum;
import cn.dream.handler.bo.CellAddressRange;
import cn.dream.handler.bo.RecordDataValidator;
import cn.dream.handler.bo.SheetData;
import cn.dream.util.ReflectionUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Validate;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.text.DateFormat;
import java.text.ParseException;
import java.util.*;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Consumer;
import java.util.function.Supplier;
import java.util.stream.Collectors;

@SuppressWarnings({"rawtypes", "unchecked"})
public abstract class AbstractExcel<T> extends WorkbookPropScope {


    /* ===========                  全局静态常量字段                       =========================  */

    /**
     * 单元格类型
     */
    private static final CellType[] CELL_TYPES = new CellType[]{
            CellType.BOOLEAN, CellType.STRING, CellType.NUMERIC, CellType.FORMULA, CellType.BLANK
    };

    /**
     * Java支持的类型列表
     */
    private static final Class<?>[] JAVA_TYPES = new Class<?>[]{
            Boolean.class, Byte.class, Short.class, Integer.class, Long.class, Float.class, Double.class, String.class, Date.class, Calendar.class
    };

    @FunctionalInterface
    interface ISetCellValue {

        default void setValue(Cell cell, Object value) throws ParseException {
            if(ObjectUtils.isEmpty(value)){
                cell.setBlank();
                return;
            }
            Validate.notNull(value,"不允许的单元格空值");
            _setValue(cell, value.toString());
        }

        void _setValue(Cell cell, String value) throws ParseException;
    }

    /**
     * 使用Java类型写入到单元格的值方法
     */
    @SuppressWarnings("Convert2MethodRef")
    private static final ISetCellValue[] JAVA_TYPE_SET_CELL_VALUE = new ISetCellValue[]{
            (cell, value) -> {
                cell.setCellValue(Boolean.parseBoolean(value));
            },
            (cell, value) -> {
                cell.setCellValue(Byte.parseByte(value));
            },
            (cell, value) -> {
                cell.setCellValue(Short.parseShort(value));
            },
            (cell, value) -> {
                cell.setCellValue(Integer.parseInt(value));
            },
            (cell, value) -> {
                cell.setCellValue(Long.parseLong(value));
            },
            (cell, value) -> {
                cell.setCellValue(Float.parseFloat(value));
            },
            (cell, value) -> {
                cell.setCellValue(Double.parseDouble(value));
            },
            (cell, value) -> {
                cell.setCellValue(value);
            },
            (cell, value) -> {
                cell.setCellValue(DateFormat.getDateTimeInstance().parse(value));
            },
            (cell, value) -> {
                Calendar calendar = Calendar.getInstance();
                calendar.setTime(DateFormat.getDateTimeInstance().parse(value));
                cell.setCellValue(calendar);
            }
    };

    /**
     * Java类型对应的映射到单元格类型的数组
     */
    private static final Short[][] JAVA_TYPE_MAPPING_CELL_TYPES = new Short[][]{
            {0}, {2}, {2}, {2}, {2}, {2}, {2}, {1, 4}
    };

    private static final CellType[] EMPTY_CELL_TYPES = new CellType[0];

    private static final String[] TYPE_STRINGS = new String[0];

    private static final Map<CellType, Class<?>> EXCEL_TYPE_MAPPING = new HashMap<>();

    protected static final String STRING_DELIMITER = "|_|";

    protected static final String EMPTY_STRING = "Null";


    /* ======          实例字段  ======================*/

    /**
     * 全局的单元格样式模具；WorkBook创建的CellStyle对象有限，需要节省使用
     */
    protected CellStyle globalCellStyle;

    protected Sheet sheet;

    /**
     * 当前对象是否是通过其他Excel转换而来；true是，false是通过本地实例的
     */
    protected boolean transfer = false;

    protected void setTransferBeTure(Object o){
        Validate.notNull(o);
        Validate.isInstanceOf(AbstractExcel.class,o);
        ReflectionUtils.setFieldValue(ReflectionUtils.getFieldByFieldName(o,"transfer").get(),o,true);
    }

    /**
     * 转换为合适的单元格类型
     *
     * @param javaType Java类型
     */
    protected static CellType[] javaTypeToCellType(Class<?> javaType) {
        int javaTypeIndex = javaTypeIndex(javaType);
        if (javaTypeIndex != -1) {
            Short[] mappingCellTypes = JAVA_TYPE_MAPPING_CELL_TYPES[javaTypeIndex];
            CellType[] cellTypes = new CellType[mappingCellTypes.length];
            for (int j = 0; j < mappingCellTypes.length; j++) {
                cellTypes[j] = CELL_TYPES[mappingCellTypes[j]];
            }
            return cellTypes;
        }
        return EMPTY_CELL_TYPES;
    }

    protected static int javaTypeIndex(Class<?> javaType) {
        for (int i = 0; i < JAVA_TYPE_MAPPING_CELL_TYPES.length; i++) {
            if (JAVA_TYPES[i] == javaType) {
                return i;
            }
        }
        return -1;
    }



    /* ===========                  实例字段                       =========================  */

    /**
     * 内嵌对象
     */
    protected boolean embeddedObject = false;

    /**
     * 本Sheet的数据
     */
    protected SheetData sheetData = null;

    /**
     * 记录数据验证的Map
     */
    private Map<Field, RecordDataValidator> recordDataValidatorMap = null;

    /**
     * 记录自动列宽的信息
     */
    private Map<Field,CellAddressRange> recordAutoColumnMap = null;

    /**
     * 记录合并单元格的信息，是指基于行自动合并的单元格信息
     */
    private Map<String, CellAddressRange> recordCellAddressRangeMap = null;


    /**
     * 缓存合并字段的列表结果，key为主字段,value为注解指定的合并单元格字段列表
     */
    protected Map<Field,List<Field>> cacheMergeFieldListMap = null;

    /**
     * 缓存合并字段的组Key名称
     */
    protected Map<Field,String> cacheMergeFieldGroupKeyMap = null;

    /**
     * 保存未执行完成的任务
     */
    protected List<Consumer<AbstractExcel<?>>> taskConsumer = new ArrayList<>();

    protected AbstractExcel(){
        recordDataValidatorMap = new HashMap<>();
        recordAutoColumnMap = new HashMap<>();
        recordCellAddressRangeMap = new LinkedHashMap<>();
        cacheMergeFieldListMap = new HashMap<>();
        cacheMergeFieldGroupKeyMap = new HashMap<>();

        if(this.workbook != null){
            // 初始化操作
            globalCellStyle = workbook.createCellStyle();
        } else {
            taskConsumer.add((abstractExcel)->{
                abstractExcel.globalCellStyle = workbook.createCellStyle();
            });
        }
    }

    public void initConsumerData(){
        taskConsumer.forEach(c -> c.accept(this));
        taskConsumer.clear();
    }

    public void oneInit(){
        cacheCellStyleMap = Collections.synchronizedMap(new HashMap<>());
    }

    /**
     * 创建一个新指向的Sheet对象
     * @param sheetName
     * @return
     */
    public abstract T newSheet(String sheetName);

    protected String validatePassReturnSafeSheetName(String sheetName){
        WorkbookUtil.validateSheetName(sheetName);
        sheetName = sheetName.trim();
        sheetName = WorkbookUtil.createSafeSheetName(sheetName);
        return sheetName;
    }

    /**
     * 如果不存在，则常见Sheet
     */
    protected Sheet createSheetIfNotExists(Workbook workbook, String sheetName) {
        String safeSheetName = validatePassReturnSafeSheetName(sheetName);
        Sheet sheet = workbook.getSheet(safeSheetName);
        if (sheet == null) {
            sheet = workbook.createSheet(safeSheetName);
        }
        return sheet;
    }

    /**
     * 设置Sheet相关的数据
     * @param dataCls
     * @param dataColl
     */
    protected <T> void setSheetData(Class<T> dataCls, List<T> dataColl){
        Field[] notStaticAndFinalFields = ReflectionUtils.getNotStaticAndFinalFields(dataCls);
        Field[] fields = Arrays.stream(notStaticAndFinalFields).filter(field -> field.isAnnotationPresent(ExcelField.class)).peek(org.springframework.util.ReflectionUtils::makeAccessible).collect(Collectors.toList()).toArray(notStaticAndFinalFields);

        List<Field> unmodifiableList = Collections.unmodifiableList(new ArrayList<>(Arrays.asList(fields)));
        this.sheetData = new SheetData<T>(dataCls, unmodifiableList, dataColl);
    }

    public SheetData getSheetData(){
        Validate.notNull(this.sheetData,"SheetData未设置,请通过 setSheetData 进行设置数据项");
        return this.sheetData;
    }

    public void createSheet(String sheetName){
        Validate.isTrue(this.sheet == null , "当前Sheet已存在对象,如需创建Sheet,请使用 #newSheet(String) 进行操作");
        this.sheet = this.createSheetIfNotExists(this.workbook,sheetName);
    }

    protected List<Field> getFields() {
        List<Field> fields = this.sheetData.getFieldList();
        return fields;
    }

    protected CellRangeAddress getCellRangeAddress(Sheet sheet, Cell cell) {
        return getCellRangeAddress(sheet,cell.getRowIndex(),cell.getColumnIndex());
    }

    /**
     * 获取指定Sheet里的指定行和列所在的合并单元格对象
     * @param sheet
     * @param rowIndex
     * @param colIndex
     * @return
     */
    protected CellRangeAddress getCellRangeAddress(Sheet sheet, final int rowIndex, final int colIndex) {
        List<CellRangeAddress> cellRangeAddressList = sheet.getMergedRegions().parallelStream().filter((cellAddresses -> cellAddresses.isInRange(rowIndex, colIndex))).collect(Collectors.toList());
        return cellRangeAddressList.size() > 0 ? cellRangeAddressList.get(0) : null;
    }

    /**
     * 获取数组中最大的值
     * @param nums
     * @return
     */
    protected int getMaxNum(int... nums) {
        int maxValue = -1;
        for (int num : nums) {
            if (num > maxValue) {
                maxValue = num;
            }
        }
        return maxValue;
    }


    /**
     * 创建模板Sheet的验证对象
     * @param selectedItems 下拉选择项的列表
     * @param cellAddressRange 应用验证的区间范围对象
     */
    protected void createDataValidator(Sheet sheet,String[] selectedItems, CellAddressRange cellAddressRange){

        CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(cellAddressRange.getFirstRow(), cellAddressRange.getLastRow(), cellAddressRange.getFirstCol(), cellAddressRange.getLastCol());

        DataValidationHelper dataValidationHelper = sheet.getDataValidationHelper();
        DataValidationConstraint explicitListConstraint = dataValidationHelper.createExplicitListConstraint(selectedItems);
        DataValidation validation = dataValidationHelper.createValidation(explicitListConstraint, cellRangeAddressList);
        validation.setShowErrorBox(true);

        if(validation instanceof XSSFDataValidation) {
            validation.setSuppressDropDownArrow(true);
            validation.setShowErrorBox(true);
        }else{
            validation.setSuppressDropDownArrow(false);
        }
        sheet.addValidationData(validation);
    }



    /**
     * 获取合并单元格的组名
     * @param o
     * @param field
     * @param sheetData
     * @return
     */
    protected String doGetGroupName(Object o,Field field,SheetData sheetData) throws IllegalAccessException {
        Object value = field.get(o);
        if(ObjectUtils.isEmpty(value)){
            return null;
        }
        return field.getName().concat(STRING_DELIMITER).concat(value.toString());

    }

    private String getMergeCellGroupName(Object o, Field field){
        if(cacheMergeFieldGroupKeyMap.containsKey(field)){
            return cacheMergeFieldGroupKeyMap.get(field);
        }

        ExcelField fieldAnnotation = field.getAnnotation(ExcelField.class);
        SheetData sheetData = this.sheetData;
        String doGetGroupName = null;
        try {
            doGetGroupName = doGetGroupName(o, field, sheetData);

            if(fieldAnnotation.mergeFields().length > 0){
                List<Field> groupKeyFieldList = null;
                if(!cacheMergeFieldListMap.containsKey(field)){
                    groupKeyFieldList = cacheMergeFieldListMap.computeIfAbsent(field, field1 -> {
                        MergeField[] mergeFields = fieldAnnotation.mergeFields();
                        Set<String> fieldSet = Arrays.stream(mergeFields).sorted(Comparator.comparingInt(MergeField::order)).map(MergeField::fieldName).collect(Collectors.toSet());
                        List<Field> fieldList = getFields();
                        return fieldList.stream().filter(f -> fieldSet.contains(f.getName())).collect(Collectors.toList());
                    });
                }

                String groupKey = groupKeyFieldList.stream().map(f -> {
                    try {
                        return f.get(o);
                    } catch (IllegalAccessException e) {
                        return EMPTY_STRING;
                    }
                }).filter(Objects::nonNull).map(String::valueOf).collect(Collectors.joining(STRING_DELIMITER));

                doGetGroupName += STRING_DELIMITER +  groupKey;
            }
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return doGetGroupName;
    }


    protected void processAndNoticeCls(Workbook workbook,Object o, Field field,Supplier<Cell> toCellSupplier,HandlerTypeEnum handlerTypeEnum) {
        processAndNoticeCls(workbook,o,field,() -> null,toCellSupplier, handlerTypeEnum);
    }

    /**
     * 后续处理的流程和通知Cls的回调
     */
    protected void processAndNoticeCls(Workbook workbook,Object o, Field field, Supplier<Cell> fromCellSupplier, Supplier<Cell> toCellSupplier, HandlerTypeEnum handlerTypeEnum) {

        Validate.notNull(handlerTypeEnum);
        Validate.notNull(field);
        Validate.notNull(toCellSupplier);

        if(HandlerTypeEnum.HEADER == handlerTypeEnum){

        }else if(HandlerTypeEnum.BODY == handlerTypeEnum){
            Validate.notNull(o);
        }

        ExcelField fieldAnnotation = field.getAnnotation(ExcelField.class);

        // 校验是否包含此字段
        DefaultApplyAnnoHandler defaultApplyAnnoHandler = ReflectionUtils.newInstance(fieldAnnotation.applyCls());
        if (fieldAnnotation.apply() && defaultApplyAnnoHandler.apply(fieldAnnotation, DefaultApplyAnnoHandler.Type.EXPORT)) {

            Cell cell = toCellSupplier.get();

            // 设置自动列宽
            if(fieldAnnotation.autoSizeColumn()){
                recordAutoColumnMap.putIfAbsent(field,CellAddressRange.builder()
                        .firstCol(cell.getColumnIndex())
                        .lastCol(cell.getColumnIndex())
                        .build());
            }


            // 这里记录下来位置，然后放到write的时候进行设置
            // 设置此Cell可选择的值列表
            if(HandlerTypeEnum.BODY == handlerTypeEnum){


                // 转换字典表达式
                String s = fieldAnnotation.converterValueExpression();
                if(StringUtils.isNotEmpty(s) || fieldAnnotation.converterValueCls() != DefaultConverterValueAnnoHandler.class ){
                    RecordDataValidator recordDataValidator = recordDataValidatorMap.computeIfAbsent(field, field1 -> {
                        Class<? extends DefaultSelectValueListAnnoHandler> selectValueListCls = fieldAnnotation.selectValueListCls();
                        DefaultSelectValueListAnnoHandler defaultSelectValueListAnnoHandler = ReflectionUtils.newInstance(selectValueListCls);
                        List<String> parseExpression = defaultSelectValueListAnnoHandler.parseExpression(fieldAnnotation.selectValues());
                        List<String> selectValueListAnnoHandlerSelectValues = defaultSelectValueListAnnoHandler.getSelectValues(parseExpression);
                        SheetData sheetData = this.sheetData;
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


                if(fieldAnnotation.mergeCell()){
                    // 记录合并单元格的范围列表
                    String groupName = getMergeCellGroupName(o, field);
                    if(StringUtils.isNotEmpty(groupName)){
                        CellAddressRange cellAddressRange = recordCellAddressRangeMap.get(groupName);
                        if(cellAddressRange == null){
                            cellAddressRange = CellAddressRange.builder().firstCol(cell.getColumnIndex()).firstRow(cell.getRowIndex()).lastCol(cell.getColumnIndex()).build();
                            recordCellAddressRangeMap.put(groupName, cellAddressRange);
                        }
                        cellAddressRange.setLastRow(cell.getRowIndex());
                    }
                }


            }

            // 设置样式单元格
            DefaultExcelFieldStyleAnnoHandler defaultExcelFieldStyleAnnoHandler = ReflectionUtils.newInstance(fieldAnnotation.cellStyleCls());
            CellStyle globalCellStyle = getGlobalCellStyle();
            Optional<Cell> cellOptional = Optional.ofNullable(fromCellSupplier.get());
            defaultExcelFieldStyleAnnoHandler.cellStyle(cellOptional.map(Cell::getCellStyle).orElse(null),globalCellStyle,handlerTypeEnum);
            globalCellStyle = createCellStyleIfNotExists(workbook,globalCellStyle);
            cell.setCellStyle(globalCellStyle);

            Class<? extends DefaultWriteValueAnnoHandler> handlerWriteValue = fieldAnnotation.handlerWriteValue();
            DefaultWriteValueAnnoHandler writeValueAnnoHandler = ReflectionUtils.newInstance(handlerWriteValue);

            try {
                field.setAccessible(true);
                AtomicReference<Class<?>> classAtomicReference = new AtomicReference<>(field.getType());
                AtomicReference<Object> valueAtomicReference = new AtomicReference<>(null);
                // 这里判断处理的类型是不是 BODY阶段，否则，o参数是为null，取不到值的，相应的默认值也不进行赋值；原因是 HEADER和FOOTER(未来可能存在)是针对注解本身的值进行操作的
                if(HandlerTypeEnum.BODY == handlerTypeEnum){
                    valueAtomicReference.compareAndSet(null,field.get(o));
                    valueAtomicReference.compareAndSet(null,fieldAnnotation.defaultValue());

                    // 当字段有值才需要进行转换
                    if(ObjectUtils.isNotEmpty(valueAtomicReference.get())){
                        // 字典转换值
                        Class<? extends DefaultConverterValueAnnoHandler> converterValueCls = fieldAnnotation.converterValueCls();
                        DefaultConverterValueAnnoHandler defaultConverterValueAnnoHandler = ReflectionUtils.newInstance(converterValueCls);
                        Map<String, String> dictDataMap = defaultConverterValueAnnoHandler.parseExpression(fieldAnnotation.converterValueExpression(), false);
                        defaultConverterValueAnnoHandler.fillConverterValue(dictDataMap);
                        defaultConverterValueAnnoHandler.doConverterValue(dictDataMap,classAtomicReference,valueAtomicReference);
                    }

                    writeValueAnnoHandler.afterHandler(classAtomicReference, valueAtomicReference);

                    // 格式化
                    Class<? extends DefaultFormatValueAnnoHandler> formatValueCls = fieldAnnotation.formatValueCls();
                    DefaultFormatValueAnnoHandler defaultFormatValueAnnoHandler = ReflectionUtils.newInstance(formatValueCls);

                    // 格式化后的值
                    valueAtomicReference.set(defaultFormatValueAnnoHandler.formatValue(valueAtomicReference.get(), classAtomicReference.get()));
                }else if(HandlerTypeEnum.HEADER == handlerTypeEnum){
                    classAtomicReference.set(String.class);
                    valueAtomicReference.set(fieldAnnotation.name());
                }

                setCellValue(cell, classAtomicReference.get(), valueAtomicReference.get());
            } catch (IllegalAccessException e) {
                onIllegalAccessException(e);
            }
        }

    }

    protected void onIllegalAccessException(IllegalAccessException illegalAccessException) {
        illegalAccessException.printStackTrace();
    }


    protected CellStyle getGlobalCellStyle() {
        return globalCellStyle;
    }

    /**
     * 获取目标Sheet中的合并单元格范围的第一个单元格对象
     *
     * @param cellAddresses
     * @return
     */
    protected Cell getMergeRangeFirstCell(Sheet sheet,CellRangeAddress cellAddresses) {
        Row targetCurrentSheetRow = sheet.getRow(cellAddresses.getFirstRow());
        if (targetCurrentSheetRow == null) {
            targetCurrentSheetRow = sheet.createRow(cellAddresses.getFirstRow());
        }

        Cell cell = targetCurrentSheetRow.getCell(cellAddresses.getFirstColumn());
        if (cell == null) {
            cell = targetCurrentSheetRow.createCell(cellAddresses.getFirstColumn());
        }
        return cell;
    }

    /**
     * 设置合并单元格的值
     * @param sheet
     * @param cellAddresses
     * @param valueType
     * @param value
     */
    protected void setCellValue(Sheet sheet,CellRangeAddress cellAddresses, Class<?> valueType, Object value) {
        Cell mergeRangeFirstCell = getMergeRangeFirstCell(sheet,cellAddresses);
        setCellValue(mergeRangeFirstCell, valueType, value);
    }

    /**
     * 设置单元格值
     *
     * @param cell      目标单元格对象
     * @param valueType 单元格的值类型
     * @param value     值
     */
    protected void setCellValue(Cell cell, Class<?> valueType, Object value) {
        int javaTypeIndex = javaTypeIndex(valueType);

        try {
            ISetCellValue iSetCellValue = JAVA_TYPE_SET_CELL_VALUE[javaTypeIndex];
            iSetCellValue.setValue(cell, value);
        } catch (ParseException e) {
            e.printStackTrace();
        }

    }

    protected Row createRowIfNotExists(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        return row;
    }

    protected Cell createCellIfNotExists(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        return cell;
    }

    /**
     * 创建单元格范围的单元格
     */
    private void createCellOfCellRangeIfNotExists(Sheet sheet,CellRangeAddress cellAddresses) {
        for (int rowIndex = cellAddresses.getFirstRow(); rowIndex < cellAddresses.getLastRow(); rowIndex++) {
            Row row = createRowIfNotExists(sheet,rowIndex);
            for (int columnIndex = cellAddresses.getFirstColumn(); columnIndex < cellAddresses.getLastColumn(); columnIndex++) {
                createCellIfNotExists(row, columnIndex);
            }
        }
    }

    protected CellStyle createCellStyleIfNotExists(CellStyle cellStyle){
        CellStyle cellStyleIfNotExists = createCellStyleIfNotExists(this.workbook, cellStyle);
        return cellStyleIfNotExists;
    }

    /**
     * 指定的样式表不存在，则进行创建，注意：手动创建的CellStyle对象，请进行保存，尽量减少创建此CellStyle对象的数量
     *
     * @param cs 模型的单元格所属的样式
     * @return
     */
    protected CellStyle createCellStyleIfNotExists(Workbook workbook,final CellStyle cs) {
        return cacheCellStyleMap.computeIfAbsent(cs.hashCode(), hashCode -> {
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.cloneStyleFrom(cs);
            return cellStyle;
        });
    }

    /**
     * 获取Model Sheet表里的合并单元格的第一个单元格对象
     *
     * @param cellAddresses
     * @return
     */
    protected Cell getFirstCell(Sheet sheet, CellRangeAddress cellAddresses) {
        return getCell(sheet,cellAddresses.getFirstRow(), cellAddresses.getFirstColumn());
    }
    protected Cell getFirstCell(Sheet sheet, Cell cell) {
        return getCell(sheet,cell.getRowIndex(),cell.getColumnIndex());
    }


    /**
     * 获取指定行指定列的单元格对象
     *
     * @param row 行索引；从0开始
     * @param col 列索引；从0开始
     * @return
     */
    protected Cell getCell(Sheet sheet, int row, int col) {
        return sheet.getRow(row).getCell(col);
    }


    /**
     * 获取 合并 和  合并单元格的值
     * @param sheet
     * @param cell
     * @return
     */
    protected Object getMergeCellValue(Sheet sheet,Cell cell){
        CellRangeAddress cellRangeAddress = getCellRangeAddress(sheet, cell);
        Object cellValue;
        if(cellRangeAddress != null){
            Cell firstCell = getFirstCell(this.sheet, cellRangeAddress);
            cellValue = getCellValue(firstCell);
        }else {
            cellValue = getCellValue(cell);
        }
        return cellValue;
    }

    protected Object getCellValue(Sheet sheet,CellRangeAddress cellAddresses) {
        return getCellValue(getFirstCell(sheet,cellAddresses));
    }

    /**
     * 获取单元格的值，并进行相应的转换类型
     *
     * @param cell
     */
    protected Object getCellValue(Cell cell) {
        CellType cellType = cell.getCellType();

        Object value = null;
        switch (cellType) {
            case STRING:
                value = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    value = cell.getDateCellValue();
                } else {
                    value = cell.getNumericCellValue();
                }
                break;
            case BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case FORMULA:
                value = cell.getCellFormula();
                break;
            case BLANK:
                value = "";
                break;
            default:
        }

        return value;
    }


    /**
     * 将临时的集合记录数据，更改写入到WorkBook中
     */
    protected void writeData(Sheet sheet){

        Validate.notNull(recordCellAddressRangeMap);
        /**
         * 合并单元格
         */
        recordCellAddressRangeMap.values().stream().filter(c -> c.getFirstRow() != c.getLastRow() || c.getFirstCol() != c.getLastCol())
                .map(c -> new CellRangeAddress(c.getFirstRow(), c.getLastRow(), c.getFirstCol(), c.getLastCol()))
                .forEach(sheet::addMergedRegion);
        recordCellAddressRangeMap.clear();


        Validate.notNull(recordAutoColumnMap);
        /**
         * 设置自动列宽
         */
        recordAutoColumnMap.forEach((field,cellAddressRange) -> {
            sheet.autoSizeColumn(cellAddressRange.getFirstCol());
        });
        recordAutoColumnMap.clear();


        Validate.notNull(recordDataValidatorMap);
        /**
         * 数据校验项列表
         */
        recordDataValidatorMap.forEach((field,recordDataValidator) -> {
            String[] selectedItems = recordDataValidator.getSelectedItems();
            if(selectedItems != null && selectedItems.length > 0){
                createDataValidator(sheet,selectedItems,recordDataValidator.getCellAddressRange());
            }
        });
        recordDataValidatorMap.clear();
    }

    protected abstract void flushData();

    /**
     * 通过 newInstance 进行实例化的调用此方法
     * @param outputFile
     * @throws IOException
     */
    public void write(File outputFile) throws IOException {
        write(this.workbook,this.sheet,outputFile);
    }



    /**
     * 将最终的数据及存放的缓存验证对等信息一同写入到Excel中
     * @param workbook
     * @param sheet
     * @param outputFile
     * @throws IOException
     */
    protected void write(Workbook workbook, Sheet sheet, File outputFile) throws IOException {
        Validate.isTrue(!this.transfer,"转换对象不能操作此方法写入数据,请通过flushData进行写入数据");
        Validate.isTrue(!this.embeddedObject,"嵌入对象不能操作Write方法");
        writeData(sheet);
        try (FileOutputStream outputStream = new FileOutputStream(outputFile)) {
            workbook.write(outputStream);
        } finally {
            workbook.close();
        }
    }


}
