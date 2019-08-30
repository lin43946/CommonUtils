package com.bgyfw.comprehensive.config;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

@Data
//@PropertySource("classpath:/alioss.properties")
public class OssInfo {
    public static String END_POINT = "";
    public static String ACCESS_KEY_ID = "";
    public static String ACCESS_KEY_SECRET = "";
    public static String BUCKET_NAME = "";
    public static String FILE_DIR = "";
}
