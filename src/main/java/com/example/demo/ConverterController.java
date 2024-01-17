package com.example.demo;

import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping("/converter")
public class ConverterController {


    @PostMapping("/eml-to-pst")
    public String convertEmlToPst(@RequestParam("file") MultipartFile file) {
        return "Conversion successful";
    }
}

