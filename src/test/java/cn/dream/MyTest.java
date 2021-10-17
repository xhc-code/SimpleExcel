package cn.dream;

import java.text.DateFormat;
import java.text.ParseException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;

public class MyTest {

    public static void main(String[] args) throws ParseException {

        Date date = new Date();

//        SimpleDateFormat

//        DateFormat.getDateTimeInstance()

        DateFormat dateTimeInstance = DateFormat.getDateTimeInstance(DateFormat.DEFAULT, DateFormat.DEFAULT);
        System.out.println(dateTimeInstance.parse("2021-8-12 9:39:11"));

        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
        LocalDateTime parse = LocalDateTime.parse("2021-08-12 09:39:11", dateTimeFormatter);
        System.out.println(parse);


//
//        date.toInstant().atZone()
//
//
//
//        LocalDate.from();



//        Date parse = DateFormat.getDateTimeInstance().parse(date.toString());
//        System.out.println(parse);

    }

}
