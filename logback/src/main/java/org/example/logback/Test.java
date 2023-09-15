package org.example.logback;

import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;

import javax.annotation.PostConstruct;

@Slf4j
@Service
public class Test {
    @PostConstruct
    public void outPrint(){
        String s = "print log";
        System.out.println(s);
        log.debug(s);
        Integer i = null;
        try{
            i.byteValue();
        }catch (Exception e) {
            System.out.println(e.toString());
            log.error(e.toString());
        }
    }
}
