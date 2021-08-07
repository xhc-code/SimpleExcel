package cn.dream.handler.bo;

import cn.dream.anno.Excel;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import org.apache.commons.lang3.Validate;

import java.lang.reflect.Field;
import java.util.List;

@Getter
@Setter
@ToString
public class SheetData<T> {

    private final Class<T> dataCls;
    private Excel clsExcel;
    /**
     * 包含 ExcelField 注解的 字段列表
     */
    private final List<Field> fieldList;
    private List<T> dataList;

    public SheetData(Class<T> dataCls, List<Field> fieldList, List<T> dataList) {
        Validate.notNull(dataCls);
        this.dataCls = dataCls;
        this.fieldList = fieldList;
        this.dataList = dataList;

        if (this.dataCls.isAnnotationPresent(Excel.class)) {
            this.clsExcel = dataCls.getAnnotation(Excel.class);
        }
    }

}
