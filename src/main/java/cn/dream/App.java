package cn.dream;

import cn.dream.test.StudentTestEntity;
import org.apache.poi.ss.usermodel.CellStyle;
import org.springframework.boot.ApplicationRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ApplicationContext;
import org.springframework.context.annotation.Bean;

import java.io.File;
import java.io.IOException;
import java.util.Arrays;

/**
 * Hello world!
 *
 */

@SpringBootApplication
public class App 
{
    public static void main( String[] args ) throws IOException {
//        test1();

//        test2();

        SpringApplication.run(App.class,args);


    }


//    @Bean
    public ApplicationRunner applicationRunner1(ApplicationContext applicationContext){
        return (args)->{
            SpringApplication.exit(applicationContext);
        };
    }


    
}
