package cn.dream.test;

import cn.dream.handler.module.WriteExcel;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Order;
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

        initData();
    }

    private static List<StudentTestEntity> studentTestEntityList = new ArrayList<>();

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

        WriteExcel writeExcel = WriteExcel.newInstance(new XSSFWorkbook());
        writeExcel.createSheet("我是学生列表");

        writeExcel.setSheetData(StudentTestEntity.class,studentTestEntityList);

        writeExcel.generateHeader();
        writeExcel.generateBody();

        writeExcel.write(targetFile);
    }



}
