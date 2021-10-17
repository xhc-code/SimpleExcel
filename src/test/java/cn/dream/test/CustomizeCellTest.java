package cn.dream.test;

import cn.dream.handler.module.WriteExcel;
import cn.dream.test2.entity.StudentEntity;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.core.io.ClassPathResource;

import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * 自定义单元格测试
 */
@SpringBootTest
public class CustomizeCellTest {

    private static final String RESULT_PREFIX_DIRECTORY = "customizeExcelOutput";

    private static ClassPathResource classPathResource;

    private static File targetFile = null;

    private static WriteExcel writeExcel;


    private static List<StudentEntity> studentTestEntityList = new ArrayList<>();

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

    public static void initData(){
        StudentEntity studentTestEntity = new StudentEntity();
        studentTestEntity.setUid("001");
        studentTestEntity.setName("恶魔001");
        studentTestEntity.setAge(21);
        studentTestEntity.setIsPublic(1);
        studentTestEntity.setSuccess(true);
        studentTestEntityList.add(studentTestEntity);

        studentTestEntity = new StudentEntity();
        studentTestEntity.setUid("001");
        studentTestEntity.setName("恶魔002");
        studentTestEntity.setAge(23);
        studentTestEntity.setIsPublic(1);
        studentTestEntity.setSuccess(true);
        studentTestEntity.setBirthday(new Date());
        studentTestEntityList.add(studentTestEntity);

        studentTestEntity = new StudentEntity();
        studentTestEntity.setUid("002");
        studentTestEntity.setName("恶魔005");
        studentTestEntity.setAge(29);
        studentTestEntity.setIsPublic(0);
        studentTestEntity.setBirthday(new Date());
        studentTestEntityList.add(studentTestEntity);
    }

    @Test
    public void test() throws ParseException {

        writeExcel.createSheet("我是Sheet恶魔");

        writeExcel.handlerCustomizeCellItem((workbook, sheet, putCellStyle,setMergeCell) -> {
            Row row = sheet.createRow(3);
            Cell cell = row.createCell(2);
            cell.setCellValue("我是魔鬼哦");

            /**
             * 此CellStyle对象可自行缓存下来，重复使用，在通过putCellStyle时，会自行缓存并返回修改过的CellStyle对象
             */
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellStyle.setFillBackgroundColor(IndexedColors.RED.getIndex());
            cell.setCellStyle(putCellStyle.cache(cellStyle));

        });

        writeExcel.setSheetData(StudentEntity.class,studentTestEntityList);

        writeExcel.generateBody();

        writeExcel.flushData();

    }

    /**
     * 最终写入到指定文件
     * @throws IOException
     */
    @AfterAll
    public static void after() throws IOException {
        writeExcel.write(targetFile);
    }


}
