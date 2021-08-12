package cn.dream.config.cs;

import cn.dream.anno.ExcelField;
import cn.dream.handler.AbstractExcel;
import cn.dream.util.DateUtils;
import org.springframework.core.convert.converter.Converter;
import org.springframework.stereotype.Component;

import java.util.Date;

@Component
public class DateToStringConverter implements Converter<Date, String> {

    /**
     * 默认的值
     */
    public static final String DATE_FORMAT = "yyyy-MM-dd HH:mm:ss";

    @Override
    public String convert(Date source) {
        System.err.println("调用了Date -> String 对象");
        String dateFormat = DATE_FORMAT;
        ExcelField localThreadExcelField = AbstractExcel.getLocalThreadExcelField();
        if(localThreadExcelField != null){
            dateFormat = localThreadExcelField.dateFormat();
        }
        return DateUtils.formatDate(source, dateFormat);
    }
}
