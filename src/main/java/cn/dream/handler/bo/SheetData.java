package cn.dream.handler.bo;

import cn.dream.anno.Excel;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import org.apache.commons.lang3.Validate;

import java.lang.reflect.Field;
import java.util.Collection;
import java.util.List;

@Getter
@Setter
@ToString
public class SheetData {

    private final Class<?> dataCls;
    private Excel clsExcel;
    /**
     * 包含 ExcelField 注解的 字段列表
     */
    private final List<Field> fieldList;
    private final List<Object> dataColl;

    public SheetData(Class<?> dataCls, List<Field> fieldList, List<Object> dataColl) {
        Validate.notNull(dataCls);
        this.dataCls = dataCls;
        this.fieldList = fieldList;
        this.dataColl = dataColl;

        if (this.dataCls.isAnnotationPresent(Excel.class)) {
            this.clsExcel = dataCls.getAnnotation(Excel.class);
        }
    }

}
