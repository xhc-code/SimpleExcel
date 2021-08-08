package cn.dream.util;

import cn.dream.excep.NotInstanceClassObjectException;
import cn.dream.test.StudentTestEntity;
import cn.dream.util.anno.Feature.RequireCopy;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.Validate;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Modifier;
import java.util.*;
import java.util.function.Predicate;
import java.util.stream.Collectors;
import java.util.stream.Stream;

@SuppressWarnings({"unchecked", "ConstantConditions"})
@Slf4j
public class ReflectionUtils {

	private static final Field[] COVER_FIELDS = new Field[0];
	
	/**
	 * 全局缓存字段集合；key为类引用路径，value为对应的字段集合
	 */
	private static final Map<String,Field[]> GLOBAL_CACHE_FIELDS = new HashMap<String, Field[]>();
	
	enum KeySuffixEnum {
		NOT_STATIC_AND_FINAL(":notStaticAndFinal");
		
		private String suffixValue;
		
		private KeySuffixEnum(String suffixValue) {
			this.suffixValue = suffixValue;
		}

		public String getSuffixValue() {
			return suffixValue;
		}
		
	}
	
	public static Field[] getNotStaticAndFinalFields(Class<?> startCls) {
		return getNotStaticAndFinalFields(startCls,null,false);
	}
	
	/**
	 * 获取字段列表(排除final和static修饰符的字段)
	 * @param startCls 起始的Cls对象，结果包含此Cls的属性
	 * @param endCls 终止的Cls对象；不包含endCls代表的此对象的属性
	 * @param parentOverride 往上遍历父级属性时，是否覆盖子类的同名属性
	 * @return
	 */
	public static Field[] getNotStaticAndFinalFields(Class<?> startCls,Class<?> endCls,boolean parentOverride) {
		String key = startCls.getName().concat(KeySuffixEnum.NOT_STATIC_AND_FINAL.getSuffixValue());
		
		if(GLOBAL_CACHE_FIELDS.containsKey(key)) {
			return GLOBAL_CACHE_FIELDS.get(key);
		}else {
			synchronized (KeySuffixEnum.NOT_STATIC_AND_FINAL) {
				if(GLOBAL_CACHE_FIELDS.containsKey(key)) {
					return GLOBAL_CACHE_FIELDS.get(key);
				}else {
					Field[] fields = getFields(startCls, endCls, parentOverride, field -> {
						int modifiers = field.getModifiers();
						if(Modifier.isStatic(modifiers) || Modifier.isFinal(modifiers)) {
							return false;
						}
						return true;
					});
					GLOBAL_CACHE_FIELDS.put(key, fields);
					return fields; 
				}
				
			}
		}
	}
	
	/**
	 * 获取类的所有字段，包含私有的字段
	 * @param startCls 起始的Cls对象，结果包含此Cls的属性
	 * @param endCls 终止的Cls对象；不包含endCls代表的此对象的属性
	 * @param parentOverride 往上遍历父级属性时，是否覆盖子类的同名属性
	 * @param filter 过滤器，返回的结果是否包含此字段Field
	 * @return 符合filter条件的Field列表
	 */
	public static Field[] getFields(Class<?> startCls,Class<?> endCls,boolean parentOverride,Predicate<Field> filter) {
		if(startCls == null) {
			throw new IllegalArgumentException("startCls参数不能为null");
		}
		
		Map<String,Field> fieldMap = new LinkedHashMap<String, Field>();
		  
		if(startCls == endCls) {
			return COVER_FIELDS;
		}
		
		if(endCls == null) {
			endCls = Object.class;
		}
		
		Class<?> tempCls = startCls;
		do {
			Field[] declaredFields = tempCls.getDeclaredFields();
			for (int i = 0; i < declaredFields.length; i++) {
				Field field = declaredFields[i];
				if(parentOverride || !fieldMap.containsKey(field.getName())) {
					if(filter.test(field)) {
						fieldMap.put(field.getName(), field);
					}
				}
			}
			
			tempCls = tempCls.getSuperclass();
		}while(tempCls != null && tempCls != endCls);

		return fieldMap.values().toArray(COVER_FIELDS);
	}

	private static final Map<Class<?>,Object> CACHE_CLASS_INSTANCE_MAP = new HashMap<>();


	public static <T> T newInstance(Class<T> cls) {
		return newInstance(cls,true);
	}

	public static <T> T newInstance(Class<T> cls,boolean single){
		try {
			return (T) _newInstance(cls,null,single);
		} catch (InstantiationException | IllegalAccessException | NoSuchMethodException | InvocationTargetException e) {
			NotInstanceClassObjectException notInstanceClassObjectException = new NotInstanceClassObjectException();
			notInstanceClassObjectException.addSuppressed(e);
			throw notInstanceClassObjectException;
		}
	}

	public static <T> T newInnerInstance(Class<T> cls,Object o) {
		return newInnerInstance(cls,o,true);
	}

	public static <T> T newInnerInstance(Class<T> cls,Object o,boolean single){
		try {
			return (T) _newInstance(cls,o,single);
		} catch (InstantiationException | IllegalAccessException | NoSuchMethodException | InvocationTargetException e) {
			NotInstanceClassObjectException notInstanceClassObjectException = new NotInstanceClassObjectException();
			notInstanceClassObjectException.addSuppressed(e);
			throw notInstanceClassObjectException;
		}
	}

	/**
	 * 实例化Class对象
	 * @param cls 要实例化的Class对象
	 * @param o 如果是非静态内部类，需要提供所属类的实例
	 * @return 返回Class表示的对象
	 * @throws InstantiationException
	 * @throws IllegalAccessException
	 * @throws NoSuchMethodException
	 * @throws InvocationTargetException
	 */
	private static Object _newInstance(Class<?> cls, Object o, boolean single) throws InstantiationException, IllegalAccessException, NoSuchMethodException, InvocationTargetException {
		Object newInstance;
		if(single && CACHE_CLASS_INSTANCE_MAP.containsKey(cls)){
			log.debug("从缓存Map中取出实例化Class对象,{}",cls.getName());
			return CACHE_CLASS_INSTANCE_MAP.get(cls);
		}
		if(cls.isMemberClass()){
			Constructor<?> constructor;
			if(Modifier.isStatic(cls.getModifiers())){
				constructor = cls.getConstructor();
				constructor.setAccessible(true);
				newInstance = constructor.newInstance();
			}else{
				Validate.notNull(o,"实例化非静态子类对象需要提供所属父级对象");
				constructor = cls.getConstructor(o.getClass());
				constructor.setAccessible(true);
				newInstance = constructor.newInstance(o);
			}
		}else{
			newInstance = cls.newInstance();
		}
		if(single){
			CACHE_CLASS_INSTANCE_MAP.putIfAbsent(cls, newInstance);
		}
		return newInstance;
	}


	private static final Predicate<Field> MAKE_ACCESSIBLE_PREDICATE = field -> {
		int modifiers = field.getModifiers();
		if(
				(!Modifier.isStatic(modifiers)
				|| !Modifier.isPublic(field.getDeclaringClass().getModifiers()))
				|| Modifier.isFinal(modifiers)
		){
			if(!field.isAccessible()){
				field.setAccessible(true);
			}
			return true;
		}
		return false;
	};

	public static Optional<Field> getFieldByFieldName(Object o, String fieldName){
		Field[] fields = getNotStaticAndFinalFields(o.getClass());
		return Arrays.stream(fields).filter(field -> field.getName().equals(fieldName)).findFirst();
	}

	/**
	 * 设置字段的值
	 */
	public static void setFieldValue(Field field,Object o,Object value) {
		if(MAKE_ACCESSIBLE_PREDICATE.test(field)){
			try {
				field.set(o,value);
			} catch (IllegalAccessException e) {
				log.warn("非法访问 {} 属性,通过反射设置值失败",field.getName());
			}
		}
	}

	private static final String[] EMPTY_STRINGS = new String[0];


	public static void copyProperties(Object source,Object target,String... copyProperties) {
		copyProperties(source, target, copyProperties,null,null,false);
	}

	public static void copyProperties(Object source,Object target,Class<?> editable,String... ignoreProperties) {
		copyProperties(source, target, null,editable,ignoreProperties,false);
	}

	public static void copyPropertiesByAnno(Object source,Object target,String... ignoreProperties) {
		copyProperties(source, target, null,null,ignoreProperties,true);
	}

	/**
	 * Copy source的属性的值到target中，需要保证字段名称一致
	 * @param source 源对象
	 * @param target 目标对象
	 * @param copyProperties 指定copy的字段名称列表；与editable互斥
	 * @param editable 限制更新的字段为一个类里包含的字段
	 * @param ignoreProperties 忽略的属性字段名称
	 * @param anno 是否开启注解行为;开启注解行为，会忽略 copyProperties 和 editable 属性
	 */
	private static void copyProperties(Object source,Object target,String[] copyProperties,Class<?> editable,String[] ignoreProperties,boolean anno){
		Validate.isTrue(
				(copyProperties !=null && (editable == null && ignoreProperties == null)) ||
						(copyProperties ==null && (editable != null ))
				|| anno
				,"copyProperties属性和editable为互斥属性");

		Field[] sourceFields = getNotStaticAndFinalFields(source.getClass());
		Field[] targetFields = getNotStaticAndFinalFields(target.getClass());

		// 基于指定Copy属性的方式复制对象的值
		List<String> copyPropSet = !anno ? new ArrayList<>(Arrays.asList(Optional.ofNullable(copyProperties).orElseGet(()->{
			Field[] editableFields = getNotStaticAndFinalFields(editable);
			Stream<String> stream = Arrays.stream(editableFields).map(Field::getName);
			if(ignoreProperties != null && ignoreProperties.length > 0){
				List<String> ignorePropertiesList = Arrays.asList(ignoreProperties);
				stream = stream.filter(v -> !ignorePropertiesList.contains(v));
			}
			return stream.collect(Collectors.toList()).toArray(EMPTY_STRINGS);
		}))) : Arrays.asList(targetFields).parallelStream().filter(field -> field.isAnnotationPresent(RequireCopy.class)).map(Field::getName).collect(Collectors.toList());

		Map<String, Field> sourceFieldMap = Arrays.asList(sourceFields).parallelStream().filter(field -> copyPropSet.contains(field.getName())).collect(Collectors.toConcurrentMap(Field::getName, field -> field));
		List<Field> copyFieldList = Arrays.asList(targetFields).parallelStream().filter(field -> copyPropSet.contains(field.getName())).sequential().collect(Collectors.toList());

		copyFieldList.forEach(field -> {
			String fieldName = field.getName();
			Field sourceField = sourceFieldMap.get(fieldName);

			MAKE_ACCESSIBLE_PREDICATE.test(sourceField);
			MAKE_ACCESSIBLE_PREDICATE.test(field);

			try {
				field.set(target,sourceField.get(source));
			} catch (IllegalAccessException e) {
				log.warn("找不到 {} 属性",fieldName);
			}
		});

	}

	public static void main(String[] args) {

		StudentTestEntity studentTestEntity = new StudentTestEntity();
		studentTestEntity.setAge(27);
		studentTestEntity.setName("我是恶魔");

		StudentTestEntity studentTestEntity2 = new StudentTestEntity();
		copyPropertiesByAnno(studentTestEntity,studentTestEntity2);

		System.out.println(studentTestEntity);
		System.out.println(studentTestEntity2);
	}



}
