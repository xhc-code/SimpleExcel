package cn.dream.anno.handler.excelfield;


import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Validate;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class DefaultSelectValueListAnnoHandler {

	static final String DELIMITER = ",";

	public List<String> parseExpression(String expression) {
		ArrayList<String> selectItemList = new ArrayList<>();
		if(StringUtils.isBlank(expression)){
			return selectItemList;
		}

		String[] split = expression.split(DELIMITER);
		selectItemList.addAll(Arrays.asList(split));
		return selectItemList;
	}

	public List<String> getSelectValues(List<String> selectItemList) {
		Validate.notNull(selectItemList);
		return selectItemList;
	}
	
}
