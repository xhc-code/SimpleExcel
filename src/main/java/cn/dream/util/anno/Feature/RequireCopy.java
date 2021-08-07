package cn.dream.util.anno.Feature;

import java.lang.annotation.*;

/**
 * 标识字段是否需要Copy;放置在目标对象的所需Copy值的字段上
 */
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Target(ElementType.FIELD)
public @interface RequireCopy {
}
