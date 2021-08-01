package cn.dream.anno.mark;

import java.lang.annotation.*;

/**
 * 未来使用的字段、类、或属性
 */
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Target({ElementType.FIELD,ElementType.TYPE,ElementType.METHOD})
public @interface FutureUse {
}
