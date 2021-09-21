package cn.dream.handler.bo;

import cn.dream.anno.Excel;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.lang.reflect.Field;
import java.util.Collections;
import java.util.List;

@Getter
@Setter
@ToString
public class SheetData<T> {

    private final Class<T> dataCls;
    private Excel excelAnno;
    /**
     * 包含 ExcelField 注解的 字段列表
     */
    private final List<Field> fieldList;
    private List<T> dataList;

    public SheetData(Class<T> dataCls, List<Field> fieldList, List<T> dataList) {
        this.dataCls = dataCls;
        this.fieldList = fieldList;
        this.dataList = dataList;

        if (dataCls != null && this.dataCls.isAnnotationPresent(Excel.class)) {
            this.excelAnno = dataCls.getAnnotation(Excel.class);
        }else{
            this.excelAnno = EmptyExcelAnno.class.getAnnotation(Excel.class);
        }
    }


    private static final SheetData DEFAULT_SHEET_DATE;

    static {
        DEFAULT_SHEET_DATE = new SheetData(null, Collections.emptyList(), Collections.emptyList());
    }

    public static <T> SheetData<T> getDefault(){
        return DEFAULT_SHEET_DATE;
    }

    @Excel(name = "Default")
    interface EmptyExcelAnno{

    }



}
