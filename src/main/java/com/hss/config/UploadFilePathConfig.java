package com.hss.config;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.config.annotation.ResourceHandlerRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

@Configuration
public class UploadFilePathConfig implements WebMvcConfigurer {

    @Value("${spring.servlet.multipart.location}")
    private String location;

    @Value("${file.staticAccessPath}")
    private String staticAccessPath;

    @Override
    public void addResourceHandlers(ResourceHandlerRegistry registry) {
        registry.addResourceHandler(staticAccessPath).addResourceLocations("file:" + location);
    }

}
