package cn.dream.test2;

import cn.dream.handler.module.WriteExcel;
import cn.dream.test.entity.MergeStudentInfoEntity;
import cn.dream.util.DateUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.*;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.core.io.ClassPathResource;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @author xiaohuichao
 * @createdDate 2021/9/27 22:22
 */
@SpringBootTest
@TestMethodOrder(value = MethodOrderer.OrderAnnotation.class)
public class MergeCellAndRead02Test {

    private static File targetDire = null;

    @BeforeAll
    public static void init() throws IOException {
        // Template目录
        ClassPathResource template = new ClassPathResource("template");
        File file = new File(template.getFile(), "自定义表头并写入数据");
        boolean b = !file.exists() && file.mkdirs();
        targetDire = file;

        initData();
    }


    private static List<MergeStudentInfoEntity> studentTestEntityList = new ArrayList<>();

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



    /**
     * 手动创建合并单元格表头并写入数据
     */
    @Test
    public void test1() throws IOException {

        WriteExcel writeExcel = WriteExcel.newInstance(new XSSFWorkbook());

        writeExcel.createSheet("我是测试");


        writeExcel.setSheetData(MergeStudentInfoEntity.class,studentTestEntityList);

        writeExcel.handlerCustomizeCellItem((workbook, sheet, cacheStyle, cellHelper) -> {
            CellRangeAddress cellRangeAddress = new CellRangeAddress(1, 2, 0, 5);
            cellHelper.writeCellValue(cellRangeAddress,"我是跨列值");

            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellStyle.setFillForegroundColor(IndexedColors.GREY_80_PERCENT.getIndex());

            cellHelper.setCellStyle(cellRangeAddress,cacheStyle.cache(cellStyle));

        });
        writeExcel.generateHeader();
        writeExcel.generateBody();



        writeExcel.write(targetDire);
    }



    @AfterAll
    public static void after(){



    }

}
