package cn.dream.handler;

import cn.dream.anno.Excel;
import cn.dream.anno.ExcelField;
import cn.dream.anno.MergeField;
import cn.dream.anno.handler.DefaultExcelNameAnnoHandler;
import cn.dream.handler.bo.CellAddressRange;
import cn.dream.handler.bo.RecordDataValidator;
import cn.dream.handler.bo.SheetData;
import cn.dream.util.ReflectionUtils;
import cn.dream.util.ValueTypeUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.ObjectUtils;
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
import java.util.function.Consumer;
import java.util.stream.Collectors;

/**
 *
 * @param <T> 创建实例返回的对象的值
 */
@Slf4j
@SuppressWarnings({"rawtypes", "unchecked"})
public abstract class AbstractExcel<T> extends WorkbookPropScope {

    protected static final Field[] EMPTY_FIELDS = new Field[0];

    protected static final CellType[] EMPTY_CELL_TYPES = new CellType[0];

    protected static final String[] TYPE_STRINGS = new String[0];

    protected static final String STRING_DELIMITER = "|_|";

    protected static final String EMPTY_STRING = "Null";




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

        /**
         * 暴露出来的执行设置Cell值方法
         * @param cell
         * @param value
         * @param cellConsumer
         * @throws ParseException
         */
        default void setValue(Cell cell, Object value,Consumer<Cell> cellConsumer) throws ParseException {
            if(ObjectUtils.isEmpty(value)){
                cell.setBlank();
                return;
            }
            Validate.notNull(value,"不允许的单元格空值");
            String finalValue;

            if(value instanceof Date){
                DateFormat dateTimeInstance = DateFormat.getDateTimeInstance();
                finalValue = dateTimeInstance.format((Date)value);
//                ExcelField localThreadExcelField = getLocalThreadExcelField();
//                finalValue = DateUtils.formatDate((Date)value,localThreadExcelField.dateFormat());
                cellConsumer.accept(cell);
            }else if(value instanceof Calendar){
                DateFormat dateTimeInstance = DateFormat.getDateTimeInstance();
                finalValue = dateTimeInstance.format(((Calendar) value).getTime());
//                ExcelField localThreadExcelField = getLocalThreadExcelField();
//                finalValue = DateUtils.formatDate(Date.from(((Calendar) value).toInstant()),localThreadExcelField.dateFormat());
                cellConsumer.accept(cell);
            }else {
                finalValue = value.toString();
            }
            _setValue(cell, finalValue);
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
                DateFormat dateTimeInstance = DateFormat.getDateTimeInstance();
                cell.setCellValue(dateTimeInstance.parse(value));
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
            {0}, {2}, {2}, {2}, {2}, {2}, {2}, {1, 4}, {2},{2}
    };

    /* ======          实例字段  ======================*/

    private CellStyle defaultCellStyle;

    /**
     * 全局的单元格样式模具；WorkBook创建的CellStyle对象有限，需要节省使用
     */
    protected CellStyle globalCellStyle;

    protected Sheet sheet;

    /**
     * 当前对象是否是通过其他Excel转换而来；true是，false是通过本地实例的
     */
    protected boolean transfer = false;

    /**
     * 设置当前对象是否是通过 转化 方法进行实例的；比如 CopyExcel -> WriteExcel 就属于转化字段,此transfer就为true
     * @param o
     */
    protected void setTransferBeTure(Object o){
        Validate.notNull(o);
        Validate.isInstanceOf(AbstractExcel.class,o);
        ReflectionUtils.setFieldValue(ReflectionUtils.getFieldByFieldName(o,"transfer").get(),o,true);
    }

    /**
     * 从Java类型转换为合适的单元格类型
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
        Validate.isTrue(javaType.getName().contains("."), "请将实体类的基本类型修改为包装类型;出问题的类型：%s",javaType.getName() );
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
    protected Map<Field, RecordDataValidator> recordDataValidatorMap = null;

    /**
     * 记录自动列宽的信息
     */
    protected Map<Field,CellAddressRange> recordAutoColumnMap = null;

    /**
     * 记录合并单元格的信息，是指基于行自动合并的单元格信息
     */
    protected Map<String, CellAddressRange> recordCellAddressRangeMap = null;


    /**
     * 缓存合并字段的列表结果，key为主字段,value为注解指定的合并单元格字段列表
     */
    protected Map<Field,List<Field>> cacheMergeFieldListMap = null;

    /**
     * 缓存合并字段的组Key名称
     */
    protected Map<Field,String> cacheMergeFieldGroupKeyMap = null;


    protected Map<String, Integer> pointerLocationMergeCellMap = null;

    /**
     * 保存未执行完成的任务
     */
    protected List<Consumer<AbstractExcel<?>>> taskConsumer = new ArrayList<>();

    /**
     * 忽略应用的字段列表
     */
    protected List<String> ignoreFieldApplyList = new ArrayList<>();

    /**
     * 实例化 实例字段，给默认值
     */
    protected AbstractExcel(){
        recordDataValidatorMap = new HashMap<>();
        recordAutoColumnMap = new HashMap<>();
        recordCellAddressRangeMap = new LinkedHashMap<>();
        cacheMergeFieldListMap = new HashMap<>();
        cacheMergeFieldGroupKeyMap = new HashMap<>();
        pointerLocationMergeCellMap = new HashMap<>();

        if(this.workbook != null){
            // 初始化操作
            defaultCellStyle = workbook.createCellStyle();
            globalCellStyle = workbook.createCellStyle();
        } else {
            taskConsumer.add((abstractExcel)->{
                abstractExcel.defaultCellStyle = workbook.createCellStyle();
                abstractExcel.globalCellStyle = workbook.createCellStyle();
            });
        }
    }

    /**
     * 避免准备实例化时,WorkBook没有值，所以延迟初始化,初始化后续的消费者列表
     */
    public void initConsumerData(){
        taskConsumer.forEach(c -> c.accept(this));
        taskConsumer.clear();
    }

    /**
     * 仅在第一次实例化对象的时候调用,xxx.newInstance;如果将 cacheCellStyleMap 保持为单例,需要确保使用的都是同一个Wordbook对象
     */
    public void oneInit(){
        cacheCellStyleMap = Collections.synchronizedMap(new HashMap<>());
    }

    /**
     * 创建一个新指向的Sheet对象
     * @param sheetName
     * @return
     */
    public abstract T newSheet(String sheetName);

    protected Workbook getWorkbook(){
        // 此错误一般不会触发
        Validate.notNull(this.workbook,"当前未设置工作簿对象，请设置Workbook对象");
        return this.workbook;
    }

    protected Sheet getSheet(){
        Validate.notNull(this.sheet,"当前未设置Sheet对象,请通过相关API进行设置");
        return this.sheet;
    }

    /**
     * 验证Sheet名称是否通过,并转化为安全的SheetName返回,避免转义符等的存在
     * @param sheetName
     * @return
     */
    protected String validatePassReturnSafeSheetName(String sheetName){
        WorkbookUtil.validateSheetName(sheetName);
        sheetName = sheetName.trim();
        sheetName = WorkbookUtil.createSafeSheetName(sheetName);
        return sheetName;
    }

    /**
     * 如果不存在，则创建Sheet
     * @param workbook 工作簿对象
     * @param sheetName Sheet名称
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
     * @param dataCls 数据集合的ofType类型
     * @param dataColl 数据集合
     */
    protected <Entity> void setSheetData(Class<Entity> dataCls, List<Entity> dataColl){
        Field[] notStaticAndFinalFields = ReflectionUtils.getNotStaticAndFinalFields(dataCls);
        Field[] fields = Arrays.stream(notStaticAndFinalFields).filter(field -> field.isAnnotationPresent(ExcelField.class)).peek(org.springframework.util.ReflectionUtils::makeAccessible).collect(Collectors.toList()).toArray(EMPTY_FIELDS);

        List<Field> unmodifiableList = Collections.unmodifiableList(new ArrayList<>(Arrays.asList(fields)));
        this.sheetData = new SheetData<Entity>(dataCls, unmodifiableList, dataColl);
    }

    /**
     * 设置忽略的应用字段列表
     * @param ignoreFieldGetterMethod 忽略的Getter方法列表
     * @param <GetterMethod>
     */
    public <GetterMethod> void setIgnoreFieldApplyList(FieldNameFunction<GetterMethod> ignoreFieldGetterMethod){
        ignoreFieldGetterMethod.getFieldSupplierList().forEach(sSupplier -> ignoreFieldApplyList.add(sSupplier.toPropertyName()));
    }

    public SheetData getSheetData(){
        Validate.notNull(this.sheetData,"SheetData未设置,请通过 setSheetData 进行设置数据项");
        return this.sheetData;
    }

    public void createSheet(String sheetName){
        Validate.isTrue(this.sheet == null , "当前Sheet已存在对象,如需创建Sheet,请使用 #newSheet(String) 进行操作");
        this.sheet = this.createSheetIfNotExists(getWorkbook(),sheetName);
    }

    protected List<Field> getFields() {
        return this.sheetData.getFieldList();
    }

    /**
     * 获取指定Sheet的指定单元格是否是合并单元格对象，如果是返回合并单元格对象，否则返回null
     * @param sheet 被操作的Sheet对象
     * @param cell 单元格
     * @return 存在返回合并单元格对象，否则返回null
     */
    protected CellRangeAddress getCellRangeAddress(Sheet sheet, Cell cell) {
        return getCellRangeAddress(sheet,cell.getRowIndex(),cell.getColumnIndex());
    }

    /**
     * 获取指定Sheet里的指定行和列所在的合并单元格对象
     * @param sheet 被操作的Sheet对象
     * @param rowIndex 单元格所在的行索引
     * @param colIndex 单元格所在的列索引
     * @return
     */
    protected CellRangeAddress getCellRangeAddress(Sheet sheet, final int rowIndex, final int colIndex) {
        List<CellRangeAddress> cellRangeAddressList = sheet.getMergedRegions().parallelStream().filter((cellAddresses -> cellAddresses.isInRange(rowIndex, colIndex))).collect(Collectors.toList());
        return cellRangeAddressList.size() > 0 ? cellRangeAddressList.get(0) : null;
    }

    /**
     * 获取数组中最大的值并返回
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
     * 创建Sheet的验证数据对象
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
     * @param o 对象
     * @param field 对象的属性
     * @param sheetData 其他信息数据
     * @return
     */
    protected String doGetGroupName(Object o,Field field,SheetData sheetData) throws IllegalAccessException {
        Object value = field.get(o);
        if(ObjectUtils.isEmpty(value)){
            return null;
        }
        return field.getName().concat(STRING_DELIMITER).concat(value.toString());
    }

    /**
     * 获取合并单元格的组名称
     * @param o
     * @param field
     * @return
     */
    protected String getMergeCellGroupName(Object o, Field field) {
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
                }else{
                    groupKeyFieldList = cacheMergeFieldListMap.get(field);
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
            log.error("非法访问{}字段，导致自动生成合并单元格组名失效,请排查问题",field.getName());
        }
        return doGetGroupName;
    }



    /**
     * [仅限WriteExcel有效]
     * 当前正在处理的currentField对应的ExcelField注解对象
     */
    protected ExcelField currentHandlerFieldAnno=null;

    protected CellStyle getGlobalCellStyle() {
        return getGlobalCellStyle(null);
    }
    protected CellStyle getGlobalCellStyle(CellStyle cellStyle) {
        globalCellStyle.cloneStyleFrom(cellStyle == null ? defaultCellStyle : cellStyle);
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
     * 将当前处理的字段放到线程本地环境中,主要是通过线程上下文共享给其他操作进行使用
     *   注意：使用 {@code setLocalThreadExcelField } 之后一定要调用 {@code clearLocalThreadExcelField} 进行清除
     */
    private static final ThreadLocal<ExcelField> EXCEL_FIELD_THREAD_LOCAL = new ThreadLocal<>();
    protected static void setLocalThreadExcelField(ExcelField excelField){
        EXCEL_FIELD_THREAD_LOCAL.set(excelField);
    }
    public static ExcelField getLocalThreadExcelField(){
        return EXCEL_FIELD_THREAD_LOCAL.get();
    }
    protected static void clearLocalThreadExcelField(){
        EXCEL_FIELD_THREAD_LOCAL.remove();
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
            Object convertValue = value;
            if(ObjectUtils.isNotEmpty(value)){
                try {
                    setLocalThreadExcelField(currentHandlerFieldAnno);
                    convertValue = ValueTypeUtils.convertValueType(value, valueType);
                }finally {
                    clearLocalThreadExcelField();
                }
            }
            try {
                setLocalThreadExcelField(currentHandlerFieldAnno);
                iSetCellValue.setValue(cell, convertValue,(c -> {

                    /**
                     * currentHandlerFieldAnno 此实例字段只有在写入Excel时，才会有值
                     */
                    if(currentHandlerFieldAnno != null){
                        CellStyle cellStyle = getGlobalCellStyle(c.getCellStyle());
                        CreationHelper creationHelper = getWorkbook().getCreationHelper();
                        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(currentHandlerFieldAnno.dateFormat()));
                        c.setCellStyle(createCellStyleIfNotExists(cellStyle));
                    }

                }));
            }finally {
                clearLocalThreadExcelField();
            }

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
     * 如果不存在，则创建合并单元格范围的单元格
     */
    protected void createCellOfCellRangeIfNotExists(Sheet sheet,CellRangeAddress cellAddresses) {
        for (int rowIndex = cellAddresses.getFirstRow(); rowIndex < cellAddresses.getLastRow(); rowIndex++) {
            Row row = createRowIfNotExists(sheet,rowIndex);
            for (int columnIndex = cellAddresses.getFirstColumn(); columnIndex < cellAddresses.getLastColumn(); columnIndex++) {
                createCellIfNotExists(row, columnIndex);
            }
        }
    }

    protected CellStyle createCellStyleIfNotExists(CellStyle cellStyle){
        return createCellStyleIfNotExists(getWorkbook(), cellStyle);
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
     * 获取Sheet表里的合并单元格的第一个单元格对象
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
            Cell firstCell = getFirstCell(getSheet(), cellRangeAddress);
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
     * @param cell 单元格对象
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

    /**
     * 将临时存起来的数据写入到Workbook中
     */
    protected abstract void flushData();

    /**
     * 通过 newInstance 进行实例化的调用此方法
     * @param outputFile
     * @throws IOException
     */
    public void write(File outputFile) throws IOException {
        write(getWorkbook(),getSheet(),outputFile);
    }

    /**
     * 将最终的数据及存放的缓存验证对等信息一同写入到Excel中
     * @param workbook WorkBook工作簿对象
     * @param sheet Sheet对象
     * @param outputFile 写出的File文件目录
     * @throws IOException
     */
    protected void write(Workbook workbook, Sheet sheet, File outputFile) throws IOException {
        Validate.isTrue(!this.transfer,"转换对象不能操作此方法写入数据,请通过flushData进行写入数据");
        Validate.isTrue(!this.embeddedObject,"嵌入对象不能操作Write方法");


        // 目录不存在则创建
        if(!outputFile.exists()){
            outputFile.mkdirs();
        }

        /**
         * 生成导出的Excel文件的名称
         */
        Excel excelAnno = getSheetData().getExcelAnno();
        DefaultExcelNameAnnoHandler defaultExcelNameAnnoHandler = ReflectionUtils.newInstance(excelAnno.handlerName());
        String excelName = defaultExcelNameAnnoHandler.getName(excelAnno.name());
        outputFile = new File(outputFile,excelName.concat(".").concat(excelAnno.extendFileType().getValue()));

        // 开始写入数据
        writeData(sheet);
        try (FileOutputStream outputStream = new FileOutputStream(outputFile)) {
            workbook.write(outputStream);
        } finally {
            workbook.close();
        }
    }


}
