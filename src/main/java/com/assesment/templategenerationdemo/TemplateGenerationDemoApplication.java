package com.assesment.templategenerationdemo;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.properties.EnableConfigurationProperties;

import com.assesment.templategenerationdemo.services.WordGenerationService;

@SpringBootApplication
public class TemplateGenerationDemoApplication {

	public static void main(String[] args) {
		SpringApplication.run(TemplateGenerationDemoApplication.class, args);
		
	}

}
