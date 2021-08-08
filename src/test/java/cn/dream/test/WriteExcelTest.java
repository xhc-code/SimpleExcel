package cn.dream.test;

import cn.dream.handler.module.WriteExcel;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.core.io.ClassPathResource;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@SpringBootTest
public class WriteExcelTest {


    private static final String RESULT_PREFIX_DIRECTORY = "writeExcelOutput";

    private static ClassPathResource classPathResource;

    private static File targetFile = null;

    @BeforeAll
    public static void init() throws IOException {
        classPathResource = new ClassPathResource("template");

        File file = classPathResource.getFile();

        file = new File(file, RESULT_PREFIX_DIRECTORY);
        boolean mkdirDire = !file.exists() && file.mkdirs();

        targetFile = new File(file, "writeExcel结果.xlsx");

        writeExcel = WriteExcel.newInstance(new XSSFWorkbook());

        initData();
    }

    private static List<StudentTestEntity> studentTestEntityList = new ArrayList<>();

    static WriteExcel writeExcel = null;

    public static void initData(){
        StudentTestEntity studentTestEntity = new StudentTestEntity();
        studentTestEntity.setUid("001");
        studentTestEntity.setName("恶魔001");
        studentTestEntity.setAge(21);
        studentTestEntityList.add(studentTestEntity);

        studentTestEntity = new StudentTestEntity();
        studentTestEntity.setUid("001");
        studentTestEntity.setName("恶魔002");
        studentTestEntity.setAge(23);
        studentTestEntityList.add(studentTestEntity);

        studentTestEntity = new StudentTestEntity();
        studentTestEntity.setUid("002");
        studentTestEntity.setName("恶魔005");
        studentTestEntity.setAge(29);
        studentTestEntityList.add(studentTestEntity);
    }

    @Test
    public void test1() throws IOException, InvalidFormatException {

        writeExcel.createSheet("我是学生列表");

        writeExcel.setSheetData(StudentTestEntity.class,studentTestEntityList);

        writeExcel.generateHeader();
        writeExcel.generateBody();


    }

    @Test
    public void test2(){
        WriteExcel writeExcel = WriteExcelTest.writeExcel.newSheet("我是年纪");

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
    public static void after() throws IOException {
        writeExcel.write(targetFile);
    }

}
