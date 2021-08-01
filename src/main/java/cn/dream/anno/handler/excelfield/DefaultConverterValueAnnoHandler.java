package cn.dream.anno.handler.excelfield;


import cn.dream.anno.ExcelField;
import org.apache.commons.collections4.MapUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Validate;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.atomic.AtomicReference;

public class DefaultConverterValueAnnoHandler {


	/**
	 * 解析表达式
	 * @param expression 表达式字符串
	 * @param reverse key和value反转位置
	 * @return key为“=”号左边,value为“=”右边；如果reverse=true，这key和Value代表的值位置相反
	 */
	public Map<String,String> parseExpression(String expression,boolean reverse) {
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
	 * @return
	 */
	public void fillConverterValue(Map<String,String> dictDataMap) {
		// TODO 这里是用户可以自行往字典Map中放值
	}


	/**
	 * 转换字典值的逻辑
	 * @param dictDataMap 字段数据Map，由 {@code #getConverterValueMap} 传值
	 * @param javaTypeCls 值的预期类型
	 * @param value 引用值，可更改此值到外部的值
	 */
	public void doConverterValue(Map<String,String> dictDataMap,AtomicReference<Class<?>> javaTypeCls,AtomicReference<Object> value){

		// todo 添加个自动单映射和多映射的操作

	}

}
