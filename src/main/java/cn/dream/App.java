package cn.dream;

import cn.dream.test.StudentTestEntity;
import org.apache.poi.ss.usermodel.CellStyle;
import org.springframework.boot.ApplicationRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ApplicationContext;
import org.springframework.context.annotation.Bean;

import java.io.File;
import java.io.IOException;
import java.util.Arrays;

/**
 * Hello world!
 *
 */

@SpringBootApplication
public class App 
{
    public static void main( String[] args ) throws IOException {
//        test1();

//        test2();

        SpringApplication.run(App.class,args);


    }


//    @Bean
    public ApplicationRunner applicationRunner1(ApplicationContext applicationContext){
        return (args)->{
            test1();

            System.out.println("完成");

            SpringApplication.exit(applicationContext);
        };
    }


    public static void test2(){
        ExcelOperate excelOperate = ExcelOperate.newWorkBook(new File("F:\\simple-excel\\testModel\\模板Excel.xlsx"));
        excelOperate.createSheet("测试目标",true);

        StudentTestEntity studentTestEntity = new StudentTestEntity();

        excelOperate.setDataCls(StudentTestEntity.class, Arrays.asList(studentTestEntity));
        excelOperate.generateHeader();

        CellStyle cellStyle = excelOperate.targetCurrentSheet.getWorkbook().createCellStyle();
        CellStyle cellStyle1 = excelOperate.targetCurrentSheet.getWorkbook().createCellStyle();
        System.out.println(cellStyle1 == cellStyle);

    }
    
    public static void test1() throws IOException {
        ExcelOperate excelOperate = ExcelOperate.newWorkBook(new File("F:\\simple-excel\\testModel\\模板Excel.xlsx"));
        excelOperate.createSheet("测试目标",true);

        StudentTestEntity studentTestEntity = new StudentTestEntity();
        studentTestEntity.setUid("001");
        studentTestEntity.setName("恶魔先生557");
        studentTestEntity.setAge(25);

        excelOperate.setDataCls(StudentTestEntity.class, Arrays.asList(studentTestEntity));
        excelOperate.generateHeader();
        excelOperate.generateBody();

        excelOperate.copyRow(8,3,1,7);
        excelOperate.copyRow(9,4,1,7);
        excelOperate.copyRow(19,19,3,0);


        excelOperate.write("F:\\simple-excel\\testModel\\结果Excel.xlsx");
    }

    
}
