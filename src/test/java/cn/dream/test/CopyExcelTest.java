package cn.dream.test;

import cn.dream.handler.module.CopyExcel;
import cn.dream.handler.module.WriteExcel;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.core.io.ClassPathResource;

import java.io.File;
import java.io.IOException;

@SpringBootTest
public class CopyExcelTest {

    private static final String RESULT_PREFIX_DIRECTORY = "copyExcelOutput";

    private static ClassPathResource classPathResource;

    private static CopyExcel copyExcel;

    private static File targetFile;


    @BeforeAll
    public static void init() throws IOException {
        classPathResource = new ClassPathResource("template");

        File file = classPathResource.getFile();

        File modalFile = new File(file, "CopyExcel模板.xlsx");



        file = new File(file, RESULT_PREFIX_DIRECTORY);
        boolean mkdirDire = !file.exists() && file.mkdirs();

        targetFile = new File(file, "CopyExcel模板_结果.xlsx");


        copyExcel = CopyExcel.newInstance(WorkbookFactory.create(modalFile), new XSSFWorkbook());
    }

    @Test
    public void copyExcel1() throws IOException {

        copyExcel.setPointDataConsumer(pointData -> {
            Object value = pointData.getValue();
            if(ObjectUtils.isNotEmpty(value) && value instanceof Double){
                double parseDouble = Double.parseDouble(value.toString());
                if(parseDouble % 2 == 0){
                    pointData.setValueWithTypeCls(String.format("我是 %f 个恶魔先生",parseDouble));
                }
            }

        });
        copyExcel.toggleFromSheet(0);

        copyExcel.createSheet("我是新加");

        copyExcel.copyRow(2,0,2,0);
        copyExcel.copyRow(6,0,6,1);

        copyExcel.flushData();

    }

    @Test
    public void test2(){

        CopyExcel copyExcel1 = copyExcel.newSheet("我是第二个Sheet");

        copyExcel1.flushData();


    }


    @Test
    public void test3(){

        WriteExcel writeExcel = copyExcel.newWriteExcel();
        writeExcel.createSheet("我是copy里的写入的");

        writeExcel.handlerCustomizeCellItem((workbook, sheet, putCellStyle) -> {

            Row row = sheet.createRow(1);
            Cell cell = row.createCell(1);
            cell.setCellValue("我是第二个Sheet");

            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(putCellStyle.apply(cellStyle));

        });

        writeExcel.flushData();

    }


    @AfterAll
    public static void des() throws IOException {


        copyExcel.write(targetFile);

    }

}
