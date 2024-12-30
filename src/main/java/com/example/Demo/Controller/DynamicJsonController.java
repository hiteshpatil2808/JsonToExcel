package com.example.Demo.Controller;

import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import java.util.Map;

@RestController
public class DynamicJsonController {

    @PostMapping("/process")
    public void processDynamicJson(@RequestBody Map<String, Object> jsonData) {
        // Display column names
        System.out.println("Column Names:");
        for (String key : jsonData.keySet()) {
            System.out.println(key);
        }

        // Display row data
        System.out.println("Row Data:");
        for (Map.Entry<String, Object> entry : jsonData.entrySet()) {
            System.out.println(entry.getKey() + ": " + entry.getValue());
        }
    }
}
