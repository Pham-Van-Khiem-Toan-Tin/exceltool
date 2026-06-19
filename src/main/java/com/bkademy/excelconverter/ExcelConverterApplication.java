package com.bkademy.excelconverter;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.scheduling.annotation.EnableScheduling;

@EnableScheduling
@SpringBootApplication
public class ExcelConverterApplication {

    public static void main(String[] args) {
        SpringApplication.run(ExcelConverterApplication.class, args);
    }

}
