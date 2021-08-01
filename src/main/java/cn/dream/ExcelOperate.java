package cn.dream;

import cn.dream.anno.Excel;
import cn.dream.anno.ExcelField;
import cn.dream.anno.MergeField;
import cn.dream.anno.handler.excelfield.*;
import cn.dream.enu.HandlerTypeEnum;
import cn.dream.enu.WorkBookTypeEnum;
import cn.dream.fun.CellItem;
import cn.dream.util.ReflectionUtils;
import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Validate;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.util.Removal;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.text.DateFormat;
import java.text.ParseException;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Function;
import java.util.function.Supplier;
import java.util.stream.Collectors;

@SuppressWarnings("ConstantConditions")
@Slf4j
public class ExcelOperate {


    private static final CellType[] CELL_TYPES = new CellType[]{
            CellType.BOOLEAN, CellType.STRING, CellType.NUMERIC, CellType.FORMULA, CellType.BLANK
    };

    private static final Class<?>[] JAVA_TYPES = new Class<?>[]{
            Boolean.class, Byte.class, Short.class, Integer.class, Long.class, Float.class, Double.class, String.class, Date.class, Calendar.class
    };

    private static final Function<?, ?>[] JAVA_TYPE_CONVERTER = new Function<?, ?>[]{
            (value) -> (boolean) value,
            (value) -> (byte) value,
            (value) -> (short) value,
            (value) -> (int) value,
            (value) -> (long) value,
            (value) -> (float) value,
            (value) -> (double) value,
            (value) -> (String) value,
            (value) -> (Date) value,
            (value) -> (Calendar) value,
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

    private static final Short[][] JAVA_TYPE_MAPPING_CELL_TYPES = new Short[][]{
            {0}, {2}, {2}, {2}, {2}, {2}, {2}, {1, 4}
    };

    private static final CellType[] EMPTY_CELL_TYPES = new CellType[0];


    // 实例化
    // copy单行，需包含普通单元格和合并单元格；是否设置值进去
    // import or export
    // 结果中是否包含此列
    // key解析规则
    // append追加行
    // 获取合并单元格的值
    // Excel数据列的解析方式

    public static final ThreadLocal<ExcelOperate> THREAD_LOCAL = new ThreadLocal<>();

    {
        THREAD_LOCAL.set(this);
    }

    private static final Map<CellType, Class<?>> EXCEL_TYPE_MAPPING = new HashMap<>();



    /**
     * 转换为合适的单元格类型
     *
     * @param javaType Java类型
     */
    public static CellType[] javaTypeToCellType(Class<?> javaType) {
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

    public static int javaTypeIndex(Class<?> javaType) {
        for (int i = 0; i < JAVA_TYPE_MAPPING_CELL_TYPES.length; i++) {
            if (JAVA_TYPES[i] == javaType) {
                return i;
            }
        }
        return -1;
    }



    /**
     * 模式位置
     *  属性的值都是以0起始，与Sheet表示的行和列一致
     */
    @Getter
    @Setter
    @ToString
    @Builder
    static class CellAddressRange {

        /**
         * 首行
         */
        private int firstRow;

        /**
         * 尾行
         */
        private int lastRow;

        /**
         * 首列
         */
        private int firstCol;

        private int lastCol;

    }

    private CellRangeAddressList fromModeAddressToCellRangeAddressList(CellAddressRange cellAddressRange){
        return new CellRangeAddressList(cellAddressRange.getFirstRow(), cellAddressRange.getLastRow(), cellAddressRange.getFirstCol(), cellAddressRange.getLastCol());
    }

    private CellAddressRange headerAddress;


    /**
     * 工作簿，目标
     */
    private Workbook targetWorkbook;
    private Workbook modelWorkbook;

    /**
     * 当前使用的Sheet,目标
     *
     * @return
     */
    public Sheet targetCurrentSheet;
    private Sheet modelCurrentSheet;

    /**
     * 本Sheet里的所有合并区域列表
     */
    private List<CellRangeAddress> cellRangeAddressList;

    /**
     * 使用过的合并单元格集合
     */
    private final Set<CellRangeAddress> useCellRangeAddressSet = new HashSet<>();

    /**
     * 记录目前准备的合并单元格；key为组名，value为合并单元格对象
     */
    private final Map<String, CellAddressRange> cellRangeAddressMap = new LinkedHashMap<>();

    /**
     * 不同Sheet数据的暂存
     */
    private final Map<Sheet, SheetData> sheetDataMap = new HashMap<>();

    /**
     * 全局的单元格样式模具
     */
    private final CellStyle GLOBAL_CELL_STYLE;

    /**
     * 单元格样式缓存Map,避免创建过多的样式对象
     */
    private Map<Integer, CellStyle> cacheCellStyleMap;
    /**
     * 内嵌对象
     */
    private boolean embeddedObject = false;

    private ExcelOperate(File modelFile, WorkBookTypeEnum workBookTypeEnum) throws IOException {
        Validate.notNull(workBookTypeEnum);

        this.modelWorkbook = WorkbookFactory.create(modelFile);

        String sheetName = this.modelWorkbook.getSheetName(0);
        toggleModelSheet(sheetName);

        if (WorkBookTypeEnum.XLSX == workBookTypeEnum) {
            this.targetWorkbook = new XSSFWorkbook();
//            XSSFWorkbook xssfWorkbook = (XSSFWorkbook) this.modelWorkbook;
//            StylesTable stylesSource = xssfWorkbook.getStylesSource();
//            XSSFWorkbook targetXssf = (XSSFWorkbook) this.targetWorkbook;
//            int numCellStyles = stylesSource.getNumCellStyles();
//
//            for(int i=0;i<numCellStyles;i++){
//                XSSFCellStyle styleAt = stylesSource.getStyleAt(i);
//
//                XSSFCellStyle cellStyle = targetXssf.createCellStyle();
//                cellStyle.cloneStyleFrom(styleAt);
//                targetXssf.getStylesSource().putStyle(cellStyle);
//            }


        } else if (WorkBookTypeEnum.XLS == workBookTypeEnum) {
            this.targetWorkbook = new HSSFWorkbook();
        }

        cacheCellStyleMap = Collections.synchronizedMap(new HashMap<>());
        GLOBAL_CELL_STYLE = this.targetWorkbook.createCellStyle();
    }


    private ExcelOperate(Workbook targetWorkbook,Workbook modelWorkbook){
        this.modelWorkbook = modelWorkbook;
        this.targetWorkbook = targetWorkbook;
        GLOBAL_CELL_STYLE = this.targetWorkbook.createCellStyle();
    }

    public static ExcelOperate newWorkBook(File modelFile) {
        try {
            ExcelOperate excelOperate = new ExcelOperate(modelFile, WorkBookTypeEnum.XLSX);
            return excelOperate;
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 创建Sheet
     *
     * @param sheetName Sheet名称
     * @param use       是否创建并使用当前目标Sheet
     */
    public void createSheet(String sheetName, boolean use) {
        WorkbookUtil.validateSheetName(sheetName);
        sheetName = sheetName.trim();
        sheetName = WorkbookUtil.createSafeSheetName(sheetName);
        Sheet sheet = this.targetWorkbook.getSheet(sheetName);
        if (sheet == null) {
            sheet = targetWorkbook.createSheet(sheetName);
            if (use) {
                this.targetCurrentSheet = sheet;
            }
            return;
        }
        throw new RuntimeException("指定的Sheet已存在");
    }

    /**
     * 切换Sheet到当前使用的Sheet对象
     *
     * @param sheetName
     */
    public void toggleSheet(String sheetName) {
        Sheet sheet = this.targetWorkbook.getSheet(sheetName);
        if (sheet != null) {
            this.targetCurrentSheet = sheet;
            return;
        }
        throw new RuntimeException("Sheet 不存在异常,请先创建Sheet");
    }

    public void toggleModelSheet(String sheetName) {
        Sheet sheet = this.modelWorkbook.getSheet(sheetName);
        if (sheet != null) {
            this.modelCurrentSheet = sheet;
            cellRangeAddressList = this.modelCurrentSheet.getMergedRegions();
            return;
        }
        throw new RuntimeException("Sheet 不存在异常,请先创建Sheet");
    }


    /**
     * 创建TargetWordBook对象的Sheet对象
     * @return 返回ExcelOperate对象，此对象表示新Sheet对象的操作
     */
    public ExcelOperate newSheetName(String sheetName){
        Validate.notBlank(sheetName);
        ExcelOperate excelOperate = new ExcelOperate(this.targetWorkbook, this.modelWorkbook);
        excelOperate.createSheet(sheetName,true);
        excelOperate.cacheCellStyleMap = this.cacheCellStyleMap;
        excelOperate.embeddedObject = true;
        return excelOperate;
    }


    /**
     * 工作表上第一个逻辑行的编号（从 0 开始）或 -1 如果不存在行
     *
     * @return
     */
    private int getTargetFirstRowIndex() {
        return this.targetCurrentSheet.getFirstRowNum();
    }

    /**
     * 行数从0开始
     *
     * @return
     */
    private int getTargetLastRowIndex() {
        return this.targetCurrentSheet.getLastRowNum();
    }

    /**
     * 创建一行
     *
     * @param rowIndex 行索引，从0开始
     * @return
     */
    public Row createRow(int rowIndex) {
        return this.targetCurrentSheet.createRow(rowIndex);
    }

    public CellRangeAddress getRangeRegion(CellItem cellItem) {
        return getRangeRegion(cellItem.getRowIndex(), cellItem.getColumnIndex());
    }

    public CellRangeAddress getRangeRegion(Cell cell) {
        return getRangeRegion(cell.getRowIndex(), cell.getColumnIndex());
    }

    /**
     * 获取指定行和列所在的合并单元格对象
     *
     * @param rowIndex 行索引，从0开始
     * @param colIndex 列索引，从0开始
     * @return
     */
    public CellRangeAddress getRangeRegion(int rowIndex, int colIndex) {
        List<CellRangeAddress> cellRangeAddressList = this.cellRangeAddressList.parallelStream().filter((cellAddresses -> cellAddresses.isInRange(rowIndex, colIndex))).collect(Collectors.toList());
        return cellRangeAddressList.size() > 0 ? cellRangeAddressList.get(0) : null;
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
        Validate.notNull(this.modelWorkbook);
        Validate.notNull(this.modelCurrentSheet);
        Validate.isTrue(modelRowIndex > -1, "模型行索引必须大于0");
        Validate.isTrue(targetRowIndex > -1, "目标行索引必须大于0");

        Row row = this.modelCurrentSheet.getRow(modelRowIndex);
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
     * Copy一个单元格到指定位置(包含对合并单元格的处理)
     * 需要确保 rowNum 和 colNum 的值大于0等于0
     *
     * @param cell   单元格对象，属于Model的单元格对象
     * @param rowIndex 行索引，从0开始，目标单元格放置的位置
     * @param colIndex 列索引，从0开始，目标单元格放置的位置
     */
    public boolean copyCell(Cell cell, int rowIndex, int colIndex) {

        CellRangeAddress rangeRegion = getRangeRegion(cell);
        if (rangeRegion != null) {
            if (useCellRangeAddressSet.contains(rangeRegion)) {
                return false;
            }

            // 处理合并单元格的值
            getCellValue(getModelFirstCell(rangeRegion));

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

            this.useCellRangeAddressSet.add(rangeRegion);
            return true;
        } else {
            getCellValue(cell);
            // 不是 合并单元格 类型
            Row targetCurrentSheetRow = this.targetCurrentSheet.getRow(rowIndex);
            if (targetCurrentSheetRow == null) {
                targetCurrentSheetRow = this.targetCurrentSheet.createRow(rowIndex);
            }

            Cell targetCurrentSheetRowCell = targetCurrentSheetRow.getCell(colIndex);
            if (targetCurrentSheetRowCell == null) {
                targetCurrentSheetRowCell = targetCurrentSheetRow.createCell(colIndex);
            }
            targetCurrentSheetRowCell.setCellStyle(createCellStyleIfNotExists(cell.getCellStyle()));
            return true;
        }
    }


    @Getter
    @Setter
    private static class SheetData {
        private final Class<?> dataCls;
        private Excel clsExcel;
        /**
         * 包含 ExcelField 注解的 字段列表
         */
        private final List<Field> fieldList;
        private final Collection<?> dataColl;
        // 当前位于循环集合的索引位置
        private int index;

        public SheetData(Class<?> dataCls, List<Field> fieldList, Collection<?> dataColl) {
            Validate.notNull(dataCls);
            this.dataCls = dataCls;
            this.fieldList = fieldList;
            this.dataColl = dataColl;

            if (this.dataCls.isAnnotationPresent(Excel.class)) {
                this.clsExcel = dataCls.getAnnotation(Excel.class);
            }
        }

    }


    /**
     * 设置数据Cls和数据集合对象
     *
     * @param dataCls
     * @param dataColl
     */
    public void setDataCls(Class<?> dataCls, Collection<?> dataColl) {
        Field[] notStaticAndFinalFields = ReflectionUtils.getNotStaticAndFinalFields(dataCls);
        Field[] fields = Arrays.stream(notStaticAndFinalFields).filter(field -> field.isAnnotationPresent(ExcelField.class)).collect(Collectors.toList()).toArray(notStaticAndFinalFields);

        List<Field> unmodifiableList = Collections.unmodifiableList(new ArrayList<>(Arrays.asList(fields)));
        SheetData sheetData = new SheetData(dataCls, unmodifiableList, dataColl);
        sheetDataMap.put(this.targetCurrentSheet, sheetData);
    }


    /**
     * 获取Sheet对应的Cls的所有带ExcelField注解的字段
     *
     * @return 返回字段的不可修改的集合列表
     */
    public List<Field> getFields() {
        List<Field> fields = getSheetData().getFieldList();
        return fields;
    }

    private SheetData getSheetData() {
        Validate.notNull(sheetDataMap.get(this.targetCurrentSheet));
        SheetData sheetData = sheetDataMap.get(targetCurrentSheet);
        return sheetData;
    }

    /**
     * 获取类上的Excel注解
     *
     * @return
     */
    public Excel getExcelAnno() {
        return getSheetData().getClsExcel();
    }


    private int getMaxNum(int... nums) {
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
    private void createDataValidator(String[] selectedItems, CellAddressRange cellAddressRange){
        DataValidationHelper dataValidationHelper = this.targetCurrentSheet.getDataValidationHelper();
        DataValidationConstraint explicitListConstraint = dataValidationHelper.createExplicitListConstraint(selectedItems);
        DataValidation validation = dataValidationHelper.createValidation(explicitListConstraint, fromModeAddressToCellRangeAddressList(cellAddressRange));
        validation.setShowErrorBox(true);

        if(validation instanceof XSSFDataValidation) {
            validation.setSuppressDropDownArrow(true);
            validation.setShowErrorBox(true);
        }else{
            validation.setSuppressDropDownArrow(false);
        }
        targetCurrentSheet.addValidationData(validation);
    }

    /**
     * 生成主体Body
     */
    public void generateBody() {

        int targetLastRowIndex = getTargetLastRowIndex() + 1;

        List<Field> fieldList = getFields();
        SheetData sheetData = getSheetData();
        Collection<?> dataColl = sheetData.getDataColl();

        AtomicInteger rowIndex = new AtomicInteger(targetLastRowIndex);
        AtomicInteger columnIndex = new AtomicInteger();
        dataColl.forEach(v -> {

            sheetData.setIndex(rowIndex.get());
            columnIndex.set(0);
            for (Field field : fieldList) {
                processAndNoticeCls(v,field,() -> {
                    Row targetSheetRowIfNotExists = createTargetSheetRowIfNotExists(rowIndex.get());
                    return createCellIfNotExists(targetSheetRowIfNotExists, columnIndex.getAndIncrement());
                }, HandlerTypeEnum.BODY);
            }

            rowIndex.getAndIncrement();
        });
    }

    private static final String STRING_DELIMITER = "|_|";

    private static final String EMPTY_STRING = "Null";

    private final Map<Field,List<Field>> cacheMergeFieldListMap = new HashMap<>();


    private String getMergeCellGroupName(Object o, Field field){
        ExcelField fieldAnnotation = field.getAnnotation(ExcelField.class);

        SheetData sheetData = getSheetData();
        String doGetGroupName = null;
        try {
            doGetGroupName = doGetGroupName(o, field, sheetData.getIndex(), sheetData.getDataColl());

            if(fieldAnnotation.mergeFields().length > 0){
                List<Field> groupKeyFieldList = cacheMergeFieldListMap.computeIfAbsent(field, field1 -> {
                    MergeField[] mergeFields = fieldAnnotation.mergeFields();
                    Set<String> fieldSet = Arrays.stream(mergeFields).sorted(Comparator.comparingInt(MergeField::order)).map(MergeField::fieldName).collect(Collectors.toSet());
                    List<Field> fieldList = getFields();
                    return fieldList.stream().filter(f -> fieldSet.contains(f.getName())).collect(Collectors.toList());
                });

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

    /**
     * 获取合并单元格的组名
     * @param o
     * @param field
     * @param index 当前正在处理的索引位置
     * @param dataList
     * @return
     */
    protected String doGetGroupName(Object o,Field field,int index,Collection<?> dataList) throws IllegalAccessException {
        Object value = field.get(o);
        if(ObjectUtils.isEmpty(value)){
            return null;
        }
        return field.getName().concat(STRING_DELIMITER).concat(value.toString());

    }

    public void processAndNoticeCls(Object o, Field field,Supplier<Cell> toCellSupplier,HandlerTypeEnum handlerTypeEnum) {
        processAndNoticeCls(o,field,() -> null,toCellSupplier, handlerTypeEnum);
    }


    private static final String[] TYPE_STRINGS = new String[0];
    /**
     * 记录数据验证的Map
     */
    private final Map<Field,RecordDataValidator> recordDataValidatorMap = new HashMap<>();

    private final Map<Field,CellAddressRange> recordAutoColumnMap = new HashMap<>();

    /**
     * 记录数据验证的信息，用于当push完毕之后，使用此对象创建相应的验证
     */
    @Getter
    @Setter
    @ToString
    @Builder
    static class RecordDataValidator {

        /**
         * 范围地址
         */
        private CellAddressRange cellAddressRange;

        /**
         * 选择项数组
         */
        private String[] selectedItems;

        private HandlerTypeEnum handlerTypeEnum;

        /**
         * 数据Cls
         */
        private Class<?> dataCls;

        /**
         * 字段对象
         */
        private Field field;

        /**
         * 对象信息
         */
        private Object o;

    }


    /**
     * 后续处理的流程和通知Cls的回调
     */
    public void processAndNoticeCls(Object o, Field field, Supplier<Cell> fromCellSupplier, Supplier<Cell> toCellSupplier, HandlerTypeEnum handlerTypeEnum) {

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
                        SheetData sheetData = getSheetData();
                        return RecordDataValidator.builder()
                                .selectedItems(selectValueListAnnoHandlerSelectValues.toArray(TYPE_STRINGS))
                                .handlerTypeEnum(handlerTypeEnum)
                                .dataCls(sheetData.dataCls)
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
                        CellAddressRange cellAddressRange = cellRangeAddressMap.get(groupName);
                        if(cellAddressRange == null){
                             cellAddressRange = CellAddressRange.builder().firstCol(cell.getColumnIndex()).firstRow(cell.getRowIndex()).lastCol(cell.getColumnIndex()).build();
                            cellRangeAddressMap.put(groupName, cellAddressRange);
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
            globalCellStyle = createCellStyleIfNotExists(globalCellStyle);
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


    /**
     * 这里已经用不到了，在更高的版本，setCellType已被弃用了
     * 此方法已弃用，将在 POI 5.0 中删除。 使用显式setCellFormula(String) 、 setCellValue(...)或setBlank()来获得所需的结果。
     *
     * @param fieldTypeCls
     * @param cell
     */
    @Deprecated
    private void setCellType(Class<?> fieldTypeCls, Cell cell) {
        CellType[] cellTypes = javaTypeToCellType(fieldTypeCls);
        CellType cellType = chooseCellType(cellTypes);
        cell.setCellType(cellType);
    }

    protected void onIllegalAccessException(IllegalAccessException illegalAccessException) {
        illegalAccessException.printStackTrace();
    }

    @FunctionalInterface
    public interface CustomizeCell {

        /**
         * 使用Sheet定义单元格和合并单元格及样式
         * @param workbook
         * @param sheet
         * @param putCellStyle 给Cell设置CellStyle对象的时候，请使用这个Lambda返回的CellStyle(已进行全局缓存样式)
         */
        void customize(Workbook workbook,Sheet sheet,Function<CellStyle,CellStyle> putCellStyle);

    }


    /**
     * 自定义处理单元格
     * @param customizeCell
     */
    public void handlerCustomizeCellItem(CustomizeCell customizeCell) {
        customizeCell.customize(this.targetWorkbook,this.targetCurrentSheet, this::createCellStyleIfNotExists);
    }

    /**
     * 将根据 @ExcelField 注解的 title 生成Header值，简便使用的方式
     */
    public void generateHeader() {
        List<Field> fields = getFields();

        Excel excelAnno = getExcelAnno();

        // 将首行取出，与目标Sheet的首行对比，
        int i = excelAnno.rowIndex();
        int lastRowNum = getTargetLastRowIndex();
        Row row = createTargetSheetRowIfNotExists(getMaxNum(i, lastRowNum, 0));
        AtomicInteger columnIndexAtomic = new AtomicInteger(getMaxNum(excelAnno.columnIndex(), row.getFirstCellNum(), 0));
        for (Field field : fields) {

            processAndNoticeCls(null,field,
                    () -> createCellIfNotExists(row, columnIndexAtomic.getAndIncrement()),
                    HandlerTypeEnum.HEADER);

        }


    }

    /**
     * 从注解里获取Header值
     *
     * @return
     */
    protected String getHeaderValue(ExcelField excelField) {
        return excelField.name();
    }


    private CellStyle getGlobalCellStyle() {
        return GLOBAL_CELL_STYLE;
    }


    public CellRangeAddress cellItemToCellRangeAddress(CellItem cellItem) {
        return new CellRangeAddress(cellItem.getRowIndex(), cellItem.getRowIndex() + cellItem.getSpanRowNum(), cellItem.getColumnIndex(), cellItem.getColumnIndex() + cellItem.getSpanColumnNum());
    }

    public CellRangeAddress addCellRangeAddress(CellItem cellItem) {
        CellRangeAddress cellAddresses = cellItemToCellRangeAddress(cellItem);
        addCellRangeAddress(cellAddresses);
        return cellAddresses;
    }

    /**
     * 添加自定义合并单元格到目标Sheet中
     *
     * @param cellAddresses
     */
    public void addCellRangeAddress(CellRangeAddress cellAddresses) {
        this.targetCurrentSheet.addMergedRegion(cellAddresses);
    }

    /**
     * 获取目标Sheet中的合并单元格范围的第一个单元格对象
     *
     * @param cellAddresses
     * @return
     */
    public Cell getMergeRangeFirstCell(CellRangeAddress cellAddresses) {
        Row targetCurrentSheetRow = this.targetCurrentSheet.getRow(cellAddresses.getFirstRow());
        if (targetCurrentSheetRow == null) {
            targetCurrentSheetRow = this.targetCurrentSheet.createRow(cellAddresses.getFirstRow());
        }

        Cell cell = targetCurrentSheetRow.getCell(cellAddresses.getFirstColumn());
        if (cell == null) {
            cell = targetCurrentSheetRow.createCell(cellAddresses.getFirstColumn());
        }
        return cell;
    }

    /**
     * 设置单元格值类型
     *
     * @param cell     将要设置的Cell单元格对象
     * @param javaType Java类型
     */
    @Deprecated
    @Removal(version = "5.0")
    private void setCellValueType(Cell cell, Class<?> javaType) {
        CellType[] cellTypes = javaTypeToCellType(javaType);
        if (EMPTY_CELL_TYPES == cellTypes) {
            throw new RuntimeException("无效Java类型");
        }
        CellType cellType = chooseCellType(cellTypes);
        cell.setCellType(cellType);
    }

    /**
     * 选择单元格合适的类型
     *
     * @param cellTypes
     * @return
     */
    protected CellType chooseCellType(CellType[] cellTypes) {
        return cellTypes[0];
    }


    public void setCellValue(CellRangeAddress cellAddresses, Class<?> valueType, Object value) {
        Cell mergeRangeFirstCell = getMergeRangeFirstCell(cellAddresses);
        setCellValue(mergeRangeFirstCell, valueType, value);
    }

    /**
     * 设置单元格值
     *
     * @param cell      目标单元格对象
     * @param valueType 单元格的值类型
     * @param value     值
     */
    public void setCellValue(Cell cell, Class<?> valueType, Object value) {
        int javaTypeIndex = javaTypeIndex(valueType);

        try {
            ISetCellValue iSetCellValue = JAVA_TYPE_SET_CELL_VALUE[javaTypeIndex];
            iSetCellValue.setValue(cell, value);
        } catch (ParseException e) {
            e.printStackTrace();
        }

    }


    /**
     * 设置单元格值
     *
     * @param cellAddresses
     * @param value
     */
    @Deprecated
    public void setCellValue(CellRangeAddress cellAddresses, Object value) {
        Cell mergeRangeFirstCell = getMergeRangeFirstCell(cellAddresses);
        setCellValue(mergeRangeFirstCell, value);
    }

    @Deprecated
    public void setCellValue(Cell cell, Object o, Field field) throws IllegalAccessException {
        field.setAccessible(true);
        setCellValue(cell, field.get(o));
    }

    @Deprecated
    public void setCellValue(Cell cell, Object value) {
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case STRING:
                System.out.println(cell.getRichStringCellValue().getString());
                cell.setCellValue(String.valueOf(value));
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    System.out.println(cell.getDateCellValue());
                    cell.setCellValue(new Date((long) value));
                } else {
                    System.out.println(cell.getNumericCellValue());
                    cell.setCellValue((double) value);
                }
                break;
            case BOOLEAN:
                System.out.println(cell.getBooleanCellValue());
                cell.setCellValue((boolean) value);
                break;
            case FORMULA:
                System.out.println(cell.getCellFormula());
                cell.setCellValue((String) value);
                break;
            case BLANK:
                cell.setBlank();
                break;
            default:
                System.out.println();
        }
    }

    private Row createTargetSheetRowIfNotExists(int rowIndex) {
        return createRowIfNotExists(this.targetCurrentSheet, rowIndex);
    }

    private Cell createTargetSheetCellOfRowIfNotExists(Row row, int columnIndex) {
        return createCellIfNotExists(row, columnIndex);
    }

    private Row createRowIfNotExists(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        return row;
    }

    private Cell createCellIfNotExists(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        return cell;
    }

    /**
     * 创建单元格范围的单元格
     */
    private void createCellOfCellRangeIfNotExists(CellRangeAddress cellAddresses) {
        for (int rowIndex = cellAddresses.getFirstRow(); rowIndex < cellAddresses.getLastRow(); rowIndex++) {
            Row targetSheetRowIfNotExists = createTargetSheetRowIfNotExists(rowIndex);
            for (int columnIndex = cellAddresses.getFirstColumn(); columnIndex < cellAddresses.getLastColumn(); columnIndex++) {
                createTargetSheetCellOfRowIfNotExists(targetSheetRowIfNotExists, columnIndex);
            }
        }
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

            Row targetRow = createRowIfNotExists(this.targetCurrentSheet, targetRange.getFirstRow() + relatedRowId);

            modelColId = modelRange.getFirstColumn();
            while (modelColId <= modelRange.getLastColumn()) {

                Cell cell = this.modelCurrentSheet.getRow(modelRowId).getCell(modelColId);
                if (cell == null) {
                    continue;
                }

                int relatedColId = modelColId - modelRange.getFirstColumn();
                Cell targetCell = createCellIfNotExists(targetRow, targetRange.getFirstColumn() + relatedColId);

                cellMap.put(cell, targetCell);

                modelColId++;
            }
        }

        this.targetCurrentSheet.addMergedRegion(targetRange);


        /**
         * 这里的单元格样式，应始终使用 {@code Workbook.createCellStyle} 进行创建，默认通过 Cell.getCellStyle()获取到的CellStyle对象是全局默认的，不应该修改，会无法正常显示样式
         */
        cellMap.forEach((modelCell, targetCell) -> targetCell.setCellStyle(createCellStyleIfNotExists(modelCell.getCellStyle())));

    }

    /**
     * 指定的样式表不存在，则进行创建，注意：手动创建的CellStyle对象，请进行保存，尽量减少创建此CellStyle对象的数量
     *
     * @param cs 模型的单元格所属的样式
     * @return
     */
    private CellStyle createCellStyleIfNotExists(final CellStyle cs) {
        return cacheCellStyleMap.computeIfAbsent(cs.hashCode(), hashCode -> {
            CellStyle cellStyle = this.targetWorkbook.createCellStyle();
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
    public Cell getModelFirstCell(CellRangeAddress cellAddresses) {
        return getModelCell(cellAddresses.getFirstRow(), cellAddresses.getFirstColumn());
    }

    /**
     * 获取指定行指定列的单元格对象
     *
     * @param row 行索引；从0开始
     * @param col 列索引；从0开始
     * @return
     */
    public Cell getModelCell(int row, int col) {
        return this.modelCurrentSheet.getRow(row).getCell(col);
    }

    public Object getCellValue(CellRangeAddress cellAddresses) {
        return getCellValue(getModelFirstCell(cellAddresses));
    }

    /**
     * 获取单元格的值，并进行相应的转换类型
     *
     * @param cell
     */
    public Object getCellValue(Cell cell) {
        CellType cellType = cell.getCellType();

        switch (cellType) {
            case STRING:
                System.out.println(cell.getRichStringCellValue().getString());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    System.out.println(cell.getDateCellValue());
                } else {
                    System.out.println(cell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                System.out.println(cell.getBooleanCellValue());
                break;
            case FORMULA:
                System.out.println(cell.getCellFormula());
                break;
            case BLANK:
                break;
            default:
        }

        return null;
    }

    /**
     * 将值转换为指定类型
     *
     * @return
     */
    private Object covertType(Object value) {

        return null;
    }





    /**
     * 将临时的集合记录数据，更改写入到WorkBook中
     */
    public void writeData(){
        /**
         * 合并单元格
         */
        cellRangeAddressMap.values().stream().filter(c -> c.getFirstRow() != c.getLastRow() || c.getFirstCol() != c.getLastCol())
                .map(c -> new CellRangeAddress(c.getFirstRow(), c.getLastRow(), c.getFirstCol(), c.getLastCol()))
                .forEach(cellAddresses -> this.targetCurrentSheet.addMergedRegion(cellAddresses));
        cellRangeAddressMap.clear();

        /**
         * 设置自动列宽
         */
        recordAutoColumnMap.forEach((field,cellAddressRange) -> {
            this.targetCurrentSheet.autoSizeColumn(cellAddressRange.getFirstCol());
        });
        recordAutoColumnMap.clear();

        /**
         * 数据校验项列表
         */
        recordDataValidatorMap.forEach((field,recordDataValidator) -> {
            String[] selectedItems = recordDataValidator.getSelectedItems();
            if(selectedItems != null && selectedItems.length > 0){
                createDataValidator(selectedItems,recordDataValidator.getCellAddressRange());
            }
        });
        recordDataValidatorMap.clear();
    }

    public void write(String output) throws IOException {
        Validate.isTrue(!this.embeddedObject,"嵌入对象不能操作Write方法");

        writeData();

        try (FileOutputStream outputStream = new FileOutputStream(output)) {
            targetWorkbook.write(outputStream);
        } finally {
            targetWorkbook.close();
            modelWorkbook.close();
        }


    }

}
