package cn.dream.util;

import org.springframework.core.io.ClassPathResource;

import java.net.URL;

public class PathUtils {

    /**
     * 获取项目根目录的路径
     * @return
     */
    public static String getProjectRootPath(){
        return System.getProperty("user.dir");
    }



    public static void main(String[] args) {
        ClassPathResource classPathResource = new ClassPathResource(".");
        System.out.println("classPathResource" + classPathResource.getPath());

        URL resource = PathUtils.class.getClassLoader().getResource(".");
        System.out.println("resource" + resource.getPath());

        URL systemResource = ClassLoader.getSystemResource(".");
        System.out.println("systemResource："  + systemResource.getPath());
    }

}
