package cn.dream.util;

import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

/**
 * 测试值类型转换工具
 */
@SpringBootTest
class ValueTypeUtilsTest {


    @Test
    void convertValueType() {

        Object o = ValueTypeUtils.convertValueType(1L, Byte.class);

        System.out.println("转换成功:" + o);

    }
}