package com.assesment.TemplateGenerationDemo;

import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import com.assesment.templategenerationdemo.services.WordGenerationService;

@SpringBootTest
class TemplateGenerationDemoApplicationTests {

	@Autowired
	WordGenerationService wordGenerationService;
	@Test
	void contextLoads() {
		wordGenerationService.executeOperation();	
	}

}
