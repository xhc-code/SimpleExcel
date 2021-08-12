package cn.dream.config.cs;

import cn.dream.anno.ExcelField;
import cn.dream.handler.AbstractExcel;
import cn.dream.util.DateUtils;
import org.springframework.core.convert.converter.Converter;
import org.springframework.stereotype.Component;

import java.util.Date;

@Component
public class StringToDateConverter implements Converter<String, Date> {

    /**
     * 默认的值
     */
    public static final String DATE_FORMAT = "yyyy-MM-dd HH:mm:ss";

    @Override
    public Date convert(String source) {
        System.err.println("调用了String -> Date 对象");
        String dateFormat = DATE_FORMAT;
        ExcelField localThreadExcelField = AbstractExcel.getLocalThreadExcelField();
        if(localThreadExcelField != null){
            dateFormat = localThreadExcelField.dateFormat();
        }
        return DateUtils.parseDate(source, dateFormat);
    }
}
