package cn.dream.anno.handler.excelfield;


import cn.dream.excep.InvalidArgumentException;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Validate;

import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.atomic.AtomicReference;

public class DefaultConverterValueAnnoHandler {

	/**
	 * 默认使用逗号作为分隔符
	 */
	private static final String DELIMITER = ",";

	/**
	 * 解析表达式
	 * @param expression 表达式字符串
	 * @param reverse key和value反转位置;true则是value-key，false则是：key-value
	 * @return key为“=”号左边,value为“=”右边；如果reverse=true，这key和Value代表的值位置相反
	 */
	public Map<String,String> parseExpression(String expression) {
		HashMap<String, String> objectObjectHashMap = new HashMap<>();
		if(StringUtils.isBlank(expression)){
			return objectObjectHashMap;
		}

		String[] split = expression.split(",");
		for (String item : split) {
			String[] split1 = item.split("=");
			Validate.isTrue(split1.length == 2,"转换值表达式格式错误;");

			objectObjectHashMap.put(split1[0],split1[1]);
		}

		return objectObjectHashMap;
	}

	/**
	 * 返回的Map用于内容转换；key为需要转换的字符串(文字)，value为转换为的值
	 * @param dictDataMap 字典数据对象
	 * @param reverse key和value反转位置;true则是value-key，false则是：key-value
	 * @return
	 */
	public void fillConverterValue(Map<String,String> dictDataMap) {
		/**
		 * 这里是用户可以自行往字典Map中放值；切记，位置别放反了
		 */
	}

	/**
	 * 反转字典集合Map对象
	 * @param dictDataMap
	 * @return
	 */
	public Map<String,String> reverse(Map<String,String> dictDataMap){
		HashMap<String, String> map = new HashMap<>();
		dictDataMap.forEach((k,v) -> map.put(v,k));
		return map;
	}


	/**
	 * 转换字典值的逻辑
	 * @param dictDataMap 字段数据Map，由 {@code #getConverterValueMap} 传值
	 * @param javaTypeCls 值的预期类型,不可更改值类型,更改成功也是无效的；表明字段的类型；写入时有效,读取时无作用
	 * @param value 引用值，可更改此值到外部的值
	 */
	public void simpleMapping(Map<String,String> dictDataMap,final AtomicReference<Class<?>> javaTypeCls,AtomicReference<Object> value){
		String s = dictDataMap.get(value.get().toString());
		value.set(s);
	}

	/**
	 * 转换字典值的逻辑;多值处理
	 * @param dictDataMap 字段数据Map，由 {@code #getConverterValueMap} 传值
	 * @param javaTypeCls 值的预期类型,不可更改值类型,更改成功也是无效的；表明字段的类型；写入时有效,读取时无作用
	 * @param value 引用值，可更改此值到外部的值
	 * @param reverse key和value反转位置;true则是value-key，false则是：key-value
	 */
	public void multiMapping(Map<String,String> dictDataMap,final AtomicReference<Class<?>> javaTypeCls,AtomicReference<Object> value){
		String oString = value.get().toString();
		String[] split = oString.split(DELIMITER);
		StringBuilder stringBuilder = new StringBuilder();
		for (String item : split) {
			String s = dictDataMap.get(item);
			if(StringUtils.isEmpty(s)){
				throw new InvalidArgumentException(String.format("找不到Key为 %s 的字典项",item));
			}
			stringBuilder.append(s).append(DELIMITER);
		}
		if(stringBuilder.length() > 0){
			stringBuilder.deleteCharAt(stringBuilder.length()-1);
		}
		javaTypeCls.set(String.class);
		value.set(stringBuilder.toString());
	}

}
