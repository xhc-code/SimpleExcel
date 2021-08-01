package cn.dream.util;

import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.function.Predicate;
import java.util.stream.Collectors;

import org.apache.commons.lang3.Validate;

import cn.dream.anno.ExcelField;

public class ExcelUtils {

	private static final Field[] COVER_FIELDS = new Field[0];
	
	/**
	 * 缓存带有注解的字段集合；key为对应的实例cls，value为cls的字段列表(仅标有ExcelField注解的字段)
	 */
	private static final Map<Class<?>,Field[]> GLOBAL_CACHE_FIELDS = new HashMap<>();
	
	private static final Object SYNC_LOCK_01 = new Object();
	
	public static Field[] getFields(Class<?> cls) {
		return getFields(cls,f -> f.isAnnotationPresent(ExcelField.class));
	}
	
	/**
	 * 获取带有 {@code ExcelField} 注解的字段列表
	 * @param cls
	 * @return
	 */
	public static Field[] getFields(Class<?> cls,Predicate<Field> fieldFilter) {
		Validate.notNull(cls);
		Validate.notNull(fieldFilter);
		
		if(GLOBAL_CACHE_FIELDS.containsKey(cls)) {
			return GLOBAL_CACHE_FIELDS.get(cls);
		}else {
			synchronized (SYNC_LOCK_01) {
				if(GLOBAL_CACHE_FIELDS.containsKey(cls)) {
					return GLOBAL_CACHE_FIELDS.get(cls);
				}else {
					Field[] fields = ReflectionUtils.getNotStaticAndFinalFields(cls);
					fields = Arrays.asList(fields).stream().filter(fieldFilter).collect(Collectors.toList()).toArray(COVER_FIELDS);
					GLOBAL_CACHE_FIELDS.put(cls, fields);
					return fields;
				}
			}
		}
	}
	
	
}
