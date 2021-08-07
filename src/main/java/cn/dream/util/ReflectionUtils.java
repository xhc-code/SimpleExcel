package cn.dream.util;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Modifier;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.function.Predicate;

import cn.dream.excep.NotInstanceClassObjectException;
import cn.dream.test.TestEntity;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.Validate;

@SuppressWarnings("unchecked")
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
		}while(tempCls != null && tempCls == endCls);

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



	
	
	
	public static void main(String[] args) throws InstantiationException, IllegalAccessException, InvocationTargetException, NoSuchMethodException {

//		newInstance(A.B.class);


		boolean memberClass = A.class.isMemberClass();
		System.out.println(memberClass);

		memberClass = A.B.class.isMemberClass();
		System.out.println(memberClass);


		memberClass = ReflectionUtils.class.isMemberClass();
		System.out.println(memberClass);

		Object newInstance = _newInstance(A.B.class,new ReflectionUtils(),true);
		System.out.println(newInstance);


	}


	public static class A {
		public class B {

		}
	}


	public static void test1(){
		Field[] fields = getFields(TestEntity.class,null,false,field -> {
			int modifiers = field.getModifiers();
			if(Modifier.isStatic(modifiers) || Modifier.isFinal(modifiers)) {
				return false;
			}
			return true;
		});
		System.out.println(fields);

		System.out.println(TestEntity.class);
		System.out.println(TestEntity.class);
		System.out.println(TestEntity.class);
		System.out.println(TestEntity.class == TestEntity.class);

	}

}
