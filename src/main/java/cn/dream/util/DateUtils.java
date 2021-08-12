package cn.dream.util;

import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;

public class DateUtils {

    /**
     * 取系统默认时区ID
     */
    private static final ZoneId ZONE_ID;

    static {
        ZONE_ID = ZoneId.systemDefault();
    }

    /**
     * 将日期字符串根据日期模式，格式化为指定的对象
     * @param dateString
     * @return
     */
    public static Date parseDate(String dateString,String datePattern){
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern(datePattern);
        LocalDateTime parse = LocalDateTime.parse(dateString, dateTimeFormatter);
        Instant instant = parse.atZone(ZONE_ID).toInstant();
        return Date.from(instant);
    }

    /**
     * 将日期对象格式化为指定模式的字符串
     * @param date
     * @param datePattern
     * @return
     */
    public static String formatDate(Date date,String datePattern){
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern(datePattern);
        Instant instant = date.toInstant();
        LocalDateTime localDateTime = ZonedDateTime.ofInstant(instant, ZONE_ID).toLocalDateTime();
        return localDateTime.format(dateTimeFormatter);
    }


    public static void main(String[] args) {

        Date date = parseDate("2021-02-05 15:11:40", "yyyy-MM-dd HH:mm:ss");

        System.out.println(String.format("parseDate: %s",date));

        System.out.println(String.format("formatDate: %s",formatDate(date,"yyyy-MM-dd HH:mm:ss")));

    }

}
