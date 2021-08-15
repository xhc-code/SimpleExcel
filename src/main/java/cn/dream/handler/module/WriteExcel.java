package cn.dream.handler.module;

import cn.dream.anno.Excel;
import cn.dream.enu.HandlerTypeEnum;
import cn.dream.handler.AbstractExcel;
import cn.dream.handler.bo.SheetData;
import cn.dream.util.ReflectionUtils;
import org.apache.commons.lang3.Validate;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.util.Collection;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

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
        int lastRowNum = getSheet().getLastRowNum();
        Row row = createRowIfNotExists(getSheet(),getMaxNum(i, lastRowNum, 0));
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

        int targetLastRowIndex = getSheet().getLastRowNum() + 1;

        List<Field> fieldList = sheetData.getFieldList();
        Collection<?> dataColl = sheetData.getDataList();

        AtomicInteger rowIndex = new AtomicInteger(targetLastRowIndex);
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

            /**
             * 设置行样式单元格信息
             */
            CellStyle globalCellStyle = getGlobalCellStyle(row.getRowStyle());
            ReflectionUtils.newInstance(excelAnno.handleRowStyle()).setRowStyle(globalCellStyle,v,rowIndex.get());
            globalCellStyle = createCellStyleIfNotExists(globalCellStyle);
            row.setRowStyle(globalCellStyle);

            rowIndex.getAndIncrement();
        });
    }


    /**
     * 自定义处理单元格
     * @param iCustomizeCell
     */
    public void handlerCustomizeCellItem(ICustomizeCell iCustomizeCell) {
        Validate.notNull(getSheet(),"请设置Sheet对象");
        iCustomizeCell.customize(getWorkbook(),getSheet(), cellStyle -> this.createCellStyleIfNotExists(getWorkbook(),cellStyle));
    }

    private WriteExcel(){}

    private WriteExcel(Workbook workbook){
        super();
        this.workbook = workbook;
        initConsumerData();
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
        ReflectionUtils.copyPropertiesByAnno(this,writeExcel);
        writeExcel.initConsumerData();
        writeExcel.createSheet(sheetName);
        writeExcel.embeddedObject = true;
        return writeExcel;
    }

    public static WriteExcel newInstance(Workbook workbook) {
        WriteExcel writeExcel = new WriteExcel(workbook);
        writeExcel.oneInit();
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
