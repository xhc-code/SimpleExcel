package cn.dream.test;

import cn.dream.handler.module.ReadExcel;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.core.io.ClassPathResource;

import java.io.File;
import java.io.IOException;

@SpringBootTest
public class ReadExcelTest {

    private static final String RESULT_PREFIX_DIRECTORY = "";

    private static ClassPathResource classPathResource;

    private static File readFile = null;

    @BeforeAll
    public static void init() throws IOException {
        classPathResource = new ClassPathResource("template");

        File file = classPathResource.getFile();

        if(StringUtils.isNotEmpty(RESULT_PREFIX_DIRECTORY)){
            file = new File(file, RESULT_PREFIX_DIRECTORY);
            boolean mkdirDire = !file.exists() && file.mkdirs();
        }

        readFile = new File(file, "读取的数据模板.xlsx");
    }


    @Test
    public void test1() throws IOException, IllegalAccessException {

        ReadExcel readExcel = ReadExcel.newInstance(WorkbookFactory.create(readFile));
        readExcel.setSheetDataCls(StudentTestEntity.class);
        readExcel.toggleSheet(0);

        readExcel.readData();



    }




}
