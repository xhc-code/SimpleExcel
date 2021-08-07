package cn.dream.test;

import cn.dream.handler.module.CopyExcel;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.poi.ss.usermodel.WorkbookFactory;
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


    }

    @Test
    public void test2(){

        copyExcel.newSheet("我是第二个Sheet");

        copyExcel.flushData();


    }


    @AfterAll
    public static void des() throws IOException {


        copyExcel.write(targetFile);

    }

}
