package com.assesment.templategenerationdemo.runner;

import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.stereotype.Component;

import com.assesment.templategenerationdemo.services.WordGenerationService;

@Component
public class ExecutionRunner implements ApplicationRunner {

    WordGenerationService wordGenerationService;

    public ExecutionRunner(WordGenerationService wordGenerationService){
        this.wordGenerationService = wordGenerationService;
    }

    @Override
    public void run(ApplicationArguments args) throws Exception {
        System.out.println("Executing");
        wordGenerationService.executeOperation();
    }
    
}
