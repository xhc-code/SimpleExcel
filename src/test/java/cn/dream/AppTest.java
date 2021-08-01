package cn.dream;

import cn.dream.test.StudentTestEntity;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.io.IOException;
import java.util.Arrays;

@SpringBootTest
public class AppTest {

    private static ExcelOperate excelOperate;

    @BeforeAll
    public static void before() throws IOException {
        excelOperate = ExcelOperate.newWorkBook(new File("F:\\simple-excel\\testModel\\模板Excel.xlsx"));


    }

    /**
     * 从原Excel文件复制到目标Excel中；包含合并单元格，单元格，和创建header和数据项
     * @throws IOException
     */
    @Test
    public void test1() throws IOException {
        excelOperate.createSheet("测试目标",true);
        StudentTestEntity studentTestEntity = new StudentTestEntity();
        studentTestEntity.setUid("001");
        studentTestEntity.setName("1743bjb");
        studentTestEntity.setAge(25);

        StudentTestEntity studentTestEntity2 = new StudentTestEntity();
        studentTestEntity2.setUid("002");
        studentTestEntity2.setName("343bj754647647674hb");
        studentTestEntity2.setAge(22);
        excelOperate.setDataCls(StudentTestEntity.class, Arrays.asList(studentTestEntity,studentTestEntity2));
        excelOperate.generateHeader();
        excelOperate.generateBody();
        excelOperate.copyRow(8,3,1,7);
        excelOperate.copyRow(9,4,1,7);
        excelOperate.copyRow(19,19,3,0);
    }

    @Test
    public void test2(){

        ExcelOperate test2 = excelOperate.newSheetName("Test2");
        StudentTestEntity studentTestEntity = new StudentTestEntity();
        studentTestEntity.setUid("00122");
        studentTestEntity.setName("1743b999jb");
        studentTestEntity.setAge(29);


        StudentTestEntity studentTestEntity2 = new StudentTestEntity();
        studentTestEntity2.setUid("00122");
        studentTestEntity2.setName("我是鬼");
        studentTestEntity2.setAge(21);

        test2.setDataCls(StudentTestEntity.class, Arrays.asList(studentTestEntity,studentTestEntity2));
        test2.generateHeader();
        test2.generateBody();

        test2.writeData();




    }

    @AfterAll
    public static void done() throws IOException {
        excelOperate.write("F:\\simple-excel\\testModel\\结果Excel.xlsx");
    }

}
