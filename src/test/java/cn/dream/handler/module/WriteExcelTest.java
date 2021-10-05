package cn.dream.handler.module;

import cn.dream.handler.module.entity.StudentInfoEntity;
import cn.dream.util.DateUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.*;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.core.io.ClassPathResource;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * 写入单元格的单元测试类
 * @author xiaohuichao
 * @createdDate 2021/10/3 22:37
 */
@SpringBootTest
@TestMethodOrder(MethodOrderer.OrderAnnotation.class)
class WriteExcelTest {

    private static final List<StudentInfoEntity> studentList = new ArrayList<>();

    private static void loadData(){
        studentList.add(StudentInfoEntity.builder()
                        .id(1L)
                        .userName("张三")
                        .age((short) 25)
                        .birthday(DateUtils.parseDate("2021-02-05 15:11:40", "yyyy-MM-dd HH:mm:ss" ))
                        .createBy(84L)
                        .createName("魔鬼名称84")
                        .createDate(new Date())
                        .memberId(22L)
                        .sex('男')
                        .userId(1394L)
                .build());


        studentList.add(StudentInfoEntity.builder()
                .id(2L)
                .userName("李四")
                .age((short) 23)
                .birthday(DateUtils.parseDate("2021-02-07 15:11:40", "yyyy-MM-dd HH:mm:ss" ))
                .createBy(84L)
                .createName("魔鬼名称423")
                .createDate(new Date())
                .memberId(99L)
                .sex('男')
                .userId(1394L)
                .build());


        studentList.add(StudentInfoEntity.builder()
                .id(1183L)
                .userName("王五")
                .age((short) 19)
                .birthday(DateUtils.parseDate("2021-03-02 15:11:40", "yyyy-MM-dd HH:mm:ss" ))
                .createBy(77L)
                .createName("魔鬼名称33")
                .createDate(new Date())
                .memberId(22L)
                .sex('男')
                .userId(1394L)
                .build());


        studentList.add(StudentInfoEntity.builder()
                .id(1185L)
                .userName("小叮当")
                .age((short) 15)
                .birthday(DateUtils.parseDate("2021-05-08 15:11:40", "yyyy-MM-dd HH:mm:ss" ))
                .createBy(77L)
                .createName("魔鬼名称993")
                .createDate(new Date())
                .memberId(22L)
                .sex('女')
                .userId(1394L)
                .build());

    }

    private static File outputDire;

    /**
     * 前置操作，准备操作
     * @author xiaohuichao
     * @createDate 2021/10/4 11:59
     */
    @BeforeAll
    public static void beforeAll() throws IOException {
        loadData();

        File template = new ClassPathResource("template").getFile();

        outputDire = new File(template, "Gen_Dire");

    }

    @Test
    public void writeData() throws IOException {
        WriteExcel writeExcel = WriteExcel.newInstance(new XSSFWorkbook());
        writeExcel.createSheet("学生记录Sheet");
        writeExcel.setSheetData(StudentInfoEntity.class,studentList);
        writeExcel.generateHeader();
        writeExcel.generateBody();

        writeExcel.flushData();
        writeExcel.write(outputDire);
    }

    /**
     * 后置处理，清理资源
     * @author xiaohuichao
     * @createDate 2021/10/4 12:01
     */
    @AfterAll
    public static void afterAll(){



    }

}