package cn.dream.anno;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.ANNOTATION_TYPE;
import static java.lang.annotation.ElementType.FIELD;

/**
 * 是否是合并字段
 */
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Target(ANNOTATION_TYPE)
public @interface MergeField {

    /**
     * 合并列的值的依据
     */
    String fieldName() default "";

    int order() default 0;

}
