package cn.dream.handler.module.helper;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.Validate;
import org.apache.poi.ss.usermodel.Cell;

import java.text.DateFormat;
import java.text.ParseException;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.function.Consumer;

/**
 * @author xiaohuichao
 * @createdDate 2021/9/24 23:30
 */
@Slf4j
public class SetCellValueHelper {

    @FunctionalInterface
    public interface ISetCellValue {

        /**
         * 暴露出来的执行设置Cell值方法
         * @param cell 单元格对象
         * @param value 单元格目标值
         * @param cellConsumer 消费者处理，发挥你的想象力
         * @throws ParseException
         */
        default void setValue(Cell cell, Object value, Consumer<Cell> cellConsumer) throws ParseException {
            if(ObjectUtils.isEmpty(value)){
                cell.setCellValue("");
                return;
            }
            Validate.notNull(value,"不允许的单元格空值");
            String finalValue;

            if(value instanceof Date){
                DateFormat dateTimeInstance = DateFormat.getDateTimeInstance();
                finalValue = dateTimeInstance.format((Date)value);
//                ExcelField localThreadExcelField = getLocalThreadExcelField();
//                finalValue = DateUtils.formatDate((Date)value,localThreadExcelField.dateFormat());
                if(cellConsumer != null){
                    cellConsumer.accept(cell);
                }
            }else if(value instanceof Calendar){
                DateFormat dateTimeInstance = DateFormat.getDateTimeInstance();
                finalValue = dateTimeInstance.format(((Calendar) value).getTime());
//                ExcelField localThreadExcelField = getLocalThreadExcelField();
//                finalValue = DateUtils.formatDate(Date.from(((Calendar) value).toInstant()),localThreadExcelField.dateFormat());
                if(cellConsumer != null){
                    cellConsumer.accept(cell);
                }
            }else {
                finalValue = value.toString();
            }
            _setValue(cell, finalValue);
        }

        void _setValue(Cell cell, String value) throws ParseException;
    }

    /**
     * 使用Java类型写入到单元格的值方法
     */
    @SuppressWarnings("Convert2MethodRef")
    private static final ISetCellValue[] JAVA_TYPE_SET_CELL_VALUE = new ISetCellValue[]{
            (cell, value) -> {
                cell.setCellValue(Boolean.parseBoolean(value));
            },
            (cell, value) -> {
                cell.setCellValue(Byte.parseByte(value));
            },
            (cell, value) -> {
                cell.setCellValue(Short.parseShort(value));
            },
            (cell, value) -> {
                cell.setCellValue(Integer.parseInt(value));
            },
            (cell, value) -> {
                cell.setCellValue(Long.parseLong(value));
            },
            (cell, value) -> {
                cell.setCellValue(Float.parseFloat(value));
            },
            (cell, value) -> {
                cell.setCellValue(Double.parseDouble(value));
            },
            (cell, value) -> {
                cell.setCellValue(value);
            },
            (cell, value) -> {
                DateFormat dateTimeInstance = DateFormat.getDateTimeInstance();
                cell.setCellValue(dateTimeInstance.parse(value));
            },
            (cell, value) -> {
                Calendar calendar = Calendar.getInstance();
                calendar.setTime(DateFormat.getDateTimeInstance().parse(value));
                cell.setCellValue(calendar);
            }
    };


    private static final Map<Class<?>, ISetCellValue> SET_CELL_VALUE_MAP = new HashMap<>();

    static {

        SET_CELL_VALUE_MAP.put(Boolean.class,JAVA_TYPE_SET_CELL_VALUE[0]);
        SET_CELL_VALUE_MAP.put(Byte.class,JAVA_TYPE_SET_CELL_VALUE[1]);
        SET_CELL_VALUE_MAP.put(Short.class,JAVA_TYPE_SET_CELL_VALUE[2]);
        SET_CELL_VALUE_MAP.put(Integer.class,JAVA_TYPE_SET_CELL_VALUE[3]);
        SET_CELL_VALUE_MAP.put(Long.class,JAVA_TYPE_SET_CELL_VALUE[4]);
        SET_CELL_VALUE_MAP.put(Float.class,JAVA_TYPE_SET_CELL_VALUE[5]);
        SET_CELL_VALUE_MAP.put(Double.class,JAVA_TYPE_SET_CELL_VALUE[6]);
        SET_CELL_VALUE_MAP.put(Character.class,JAVA_TYPE_SET_CELL_VALUE[7]);
        SET_CELL_VALUE_MAP.put(String.class,JAVA_TYPE_SET_CELL_VALUE[7]);
        SET_CELL_VALUE_MAP.put(Date.class,JAVA_TYPE_SET_CELL_VALUE[8]);
        SET_CELL_VALUE_MAP.put(Calendar.class,JAVA_TYPE_SET_CELL_VALUE[9]);

    }

    /**
     * 获取设置值单元格的处理程序
     * @param javaType
     * @return
     */
    public static ISetCellValue getSetValueCell(Class<?> javaType){
        return SET_CELL_VALUE_MAP.get(javaType);
    }

}
