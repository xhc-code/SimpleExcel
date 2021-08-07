package cn.dream.handler.module;

import cn.dream.anno.Excel;
import cn.dream.enu.HandlerTypeEnum;
import cn.dream.handler.AbstractExcel;
import cn.dream.handler.WorkbookPropScope;
import cn.dream.handler.bo.SheetData;
import org.apache.commons.lang3.Validate;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.BeanUtils;

import java.lang.reflect.Field;
import java.util.Collection;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

public class WriteExcel extends AbstractExcel<WriteExcel> {

    /**
     * 将根据 @ExcelField 注解的 title 生成Header值，简便使用的方式
     */
    public void generateHeader() {
        Validate.notNull(this.sheet);

        SheetData sheetData = getSheetData();
        Excel excelAnno = sheetData.getClsExcel();

        List<Field> fields = sheetData.getFieldList();

        // 将首行取出，与目标Sheet的首行对比，
        int i = excelAnno.rowIndex();
        int lastRowNum = this.sheet.getLastRowNum();
        Row row = createRowIfNotExists(this.sheet,getMaxNum(i, lastRowNum, 0));
        AtomicInteger columnIndexAtomic = new AtomicInteger(getMaxNum(excelAnno.columnIndex(), row.getFirstCellNum(), 0));
        for (Field field : fields) {

            processAndNoticeCls(this.workbook,null,field,
                    () -> createCellIfNotExists(row, columnIndexAtomic.getAndIncrement()),
                    HandlerTypeEnum.HEADER);

        }

    }

    @Override
    public void setSheetData(Class<?> dataCls, Collection<?> dataColl) {
        super.setSheetData(dataCls, dataColl);
    }

    /**
     * 生成主体Body
     */
    public void generateBody() {
        Validate.notNull(this.sheet);

        final SheetData sheetData = getSheetData();

        int targetLastRowIndex = this.sheet.getLastRowNum() + 1;

        List<Field> fieldList = sheetData.getFieldList();
        Collection<?> dataColl = sheetData.getDataColl();

        AtomicInteger rowIndex = new AtomicInteger(targetLastRowIndex);
        AtomicInteger columnIndex = new AtomicInteger();
        dataColl.forEach(v -> {
            columnIndex.set(0);

            for (Field field : fieldList) {
                processAndNoticeCls(this.workbook,v,field,() -> {
                    Row targetSheetRowIfNotExists = createRowIfNotExists(this.sheet,rowIndex.get());
                    return createCellIfNotExists(targetSheetRowIfNotExists, columnIndex.getAndIncrement());
                }, HandlerTypeEnum.BODY);
            }

            rowIndex.getAndIncrement();
        });
    }


    /**
     * 自定义处理单元格
     * @param iCustomizeCell
     */
    public void handlerCustomizeCellItem(ICustomizeCell iCustomizeCell) {
        iCustomizeCell.customize(this.workbook,this.sheet, cellStyle -> this.createCellStyleIfNotExists(this.workbook,cellStyle));
    }

    private WriteExcel(){
        globalCellStyle = workbook.createCellStyle();
    }

    private WriteExcel(Workbook workbook){
        super(true);
        this.workbook = workbook;

        // 初始化操作
        globalCellStyle = workbook.createCellStyle();
    }

    @Override
    public WriteExcel newSheet(String sheetName) {
        WriteExcel writeExcel = new WriteExcel();
        BeanUtils.copyProperties(this,writeExcel, WorkbookPropScope.class);
        writeExcel.createSheet(sheetName);
        return writeExcel;
    }

    public static WriteExcel newInstance(Workbook workbook) {
        WriteExcel writeExcel = new WriteExcel(workbook);
        return writeExcel;
    }


}
