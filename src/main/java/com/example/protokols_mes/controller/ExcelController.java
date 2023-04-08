package com.example.protokols_mes.controller;

import com.example.protokols_mes.entity.Protokol;
import com.example.protokols_mes.excel.ExcelWriter;
import com.example.protokols_mes.service.ProtokolService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;

import java.io.IOException;

@Controller
public class ExcelController {

    @Autowired
    private ExcelWriter excelWriter;

    @Autowired
    private ProtokolService protokolService;

//    @GetMapping(value = "/excel")
//    public String excel(@RequestParam(value = "protokolId") Long protokolId) throws IOException {
//        Protokol protokol = protokolService.findById(protokolId);
//        excelWriter.excelWriteProtokol(protokol);
//        return "SUCCESS";
//    }
}
