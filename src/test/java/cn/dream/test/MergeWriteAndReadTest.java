package cn.dream.test;

import cn.dream.handler.module.ReadExcel;
import cn.dream.handler.module.WriteExcel;
import cn.dream.test.entity.MergeStudentInfoEntity;
import cn.dream.util.DateUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.*;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.core.io.ClassPathResource;

import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;

/**
 * @author xiaohuichao
 * @createdDate 2021/9/21 12:12
 */
@SpringBootTest
@TestMethodOrder(value = MethodOrderer.OrderAnnotation.class)
public class MergeWriteAndReadTest {


    private static List<MergeStudentInfoEntity> studentTestEntityList = new ArrayList<>();

    private static ClassPathResource classPathResource;

    private static File writeOutputFile;

    private static WriteExcel writeExcel;


    @BeforeAll
    public static void init() throws IOException {
        classPathResource = new ClassPathResource("template");

        File file = classPathResource.getFile();

        file = new File(file, MergeWriteAndReadTest.class.getSimpleName());
        boolean mkdirDire = !file.exists() && file.mkdirs();

        writeOutputFile = new File(file, MergeWriteAndReadTest.class.getSimpleName());

        writeExcel = WriteExcel.newInstance(new XSSFWorkbook());

        initData();
    }

    /**
     * 初始化数据
     */
    public static void initData(){
        MergeStudentInfoEntity studentTestEntity = new MergeStudentInfoEntity();
        studentTestEntity.setUid("001");
        studentTestEntity.setUserName("张三01");
        studentTestEntity.setSex('男');
        studentTestEntity.setAge(27);
        studentTestEntity.setBirthday(DateUtils.parseDate("2015-02-03 15:55:00","yyyy-MM-dd HH:mm:ss"));
        studentTestEntity.setAuditStatus(0);
        studentTestEntity.setIsPublic(1);
        studentTestEntity.setRecordDate("2017-02-03 15:55:00");
        studentTestEntity.setCreateBy(2);
        studentTestEntity.setCreateName("admin管理员");
        studentTestEntityList.add(studentTestEntity);


        studentTestEntity = new MergeStudentInfoEntity();
        studentTestEntity.setUid("002");
        studentTestEntity.setUserName("张三02");
        studentTestEntity.setSex('女');
        studentTestEntity.setAge(25);
        studentTestEntity.setBirthday(DateUtils.parseDate("2011-02-03 15:55:00","yyyy-MM-dd HH:mm:ss"));
        studentTestEntity.setAuditStatus(2);
        studentTestEntity.setIsPublic(2);
        studentTestEntity.setRecordDate("2018-02-03 15:55:00");
        studentTestEntity.setCreateBy(2);
        studentTestEntity.setCreateName("admin管理员");
        studentTestEntityList.add(studentTestEntity);


        studentTestEntity = new MergeStudentInfoEntity();
        studentTestEntity.setUid("001");
        studentTestEntity.setUserName("张三01");
        studentTestEntity.setSex('男');
        studentTestEntity.setAge(22);
        studentTestEntity.setBirthday(DateUtils.parseDate("2014-02-03 15:55:00","yyyy-MM-dd HH:mm:ss"));
        studentTestEntity.setAuditStatus(0);
        studentTestEntity.setIsPublic(1);
        studentTestEntity.setRecordDate("2019-02-03 15:55:00");
        studentTestEntity.setCreateBy(2);
        studentTestEntity.setCreateName("admin管理员");
        studentTestEntityList.add(studentTestEntity);

    }



    @Test
    @Order(0)
    public void write() throws ParseException {

        writeExcel.createSheet("我是一个测试MergeWrite的Sheet");

        writeExcel.setSheetData(MergeStudentInfoEntity.class,studentTestEntityList);

        writeExcel.handlerCustomizeCellItem((workbook, sheet, putCellStyle, cellHelper) -> {
            cellHelper.writeCellValue(new CellRangeAddress(0,2,0,0),"用户UID");

            cellHelper.writeCellValue(new CellRangeAddress(0,0,1,4),"基本信息");
            cellHelper.writeCellValue(new CellRangeAddress(1,2,1,1),"用户名称");
            cellHelper.writeCellValue(new CellRangeAddress(1,2,2,2),"用户年龄");
            cellHelper.writeCellValue(new CellRangeAddress(1,2,3,3),"用户性别");
            cellHelper.writeCellValue(new CellRangeAddress(1,2,4,4),"生日日期");

            cellHelper.writeCellValue(new CellRangeAddress(0,0,5,9),"其他信息");
            cellHelper.writeCellValue(new CellRangeAddress(1,2,5,5),"记录日期");
            cellHelper.writeCellValue(new CellRangeAddress(1,2,6,6),"创建ID");
            cellHelper.writeCellValue(new CellRangeAddress(1,2,7,7),"创建名称");
            cellHelper.writeCellValue(new CellRangeAddress(1,2,8,8),"审核状态");
            cellHelper.writeCellValue(new CellRangeAddress(1,2,9,9),"是否公开");



        });

        writeExcel.generateBody();
    }


    @Test
    @Order(1)
    public void writeComplete() throws IOException {
        writeOutputFile = writeExcel.write(writeOutputFile);
        System.out.println("准备读取----------------------------------------------");
    }

    @Test
    @Order(2)
    public void read() throws IOException, IllegalAccessException, InvalidFormatException {

        ReadExcel readExcel = ReadExcel.newInstance(WorkbookFactory.create(writeOutputFile));

        readExcel.setSheetDataCls(MergeStudentInfoEntity.class);

        readExcel.toggleSheet(0);
        readExcel.readData();

        List<MergeStudentInfoEntity> result = readExcel.getResult();
        result.forEach(System.err::println);

    }


    public void mergeWrite(){




    }


    public void mergeRead(){

    }

    /* Excel写入，读取 */

    /* 合并Excel写入，读取*/




    /**
     * 最终写入到指定文件
     * @throws IOException
     */
    @AfterAll
    public static void after() throws IOException {

    }


}
