package com.example.protokols_mes.controller;

import com.example.protokols_mes.entity.Protokol;
import com.example.protokols_mes.excel.ExcelWriter;
import com.example.protokols_mes.service.ProtokolService;
import org.apache.catalina.LifecycleState;
import org.hibernate.type.descriptor.java.ShortPrimitiveArrayJavaType;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.ModelAndView;

import java.io.IOException;
import java.util.List;

@Controller
public class ProtokolController {

    @Autowired
    private ExcelWriter excelWriter;

    @Autowired
    private ProtokolService protokolService;

    @RequestMapping({"/", "/search"})
    public String protokolPage(Protokol protokol, Model model, String keyword) {
        if (keyword != null) {
            List<Protokol> protokols = protokolService.getProtokolByName(keyword);
            model.addAttribute("protokols", protokols);
        } else {
            List<Protokol> protokols = protokolService.findAllProtokol();
            model.addAttribute("protokols", protokols);
        }
            return "protokolPage";
    }

    @GetMapping(value = "/addProtokolForm")
    public ModelAndView addProtokolForm() {
        ModelAndView modelAndView = new ModelAndView("add-protokol-form");
        Protokol newProtokol = new Protokol();
        modelAndView.addObject("protokol", newProtokol);
        return modelAndView;
    }

    @PostMapping(value = "/saveProtokol")
    public String saveProtokol(@ModelAttribute Protokol protokol) {
        this.protokolService.save(protokol);
        return "redirect:/";
    }

    @GetMapping(value = "/protokolUpdateForm")
    public ModelAndView protokolUpdateForm(@RequestParam Long protokolId) {
        ModelAndView modelAndView = new ModelAndView("add-protokol-form");
        Protokol protokol = protokolService.findById(protokolId);
        modelAndView.addObject("protokol", protokol);
        return modelAndView;
    }

    @GetMapping(value = "/excel")
    public String excel(@RequestParam(value = "protokolId") Long protokolId) throws IOException {
        Protokol protokol = protokolService.findById(protokolId);
        excelWriter.excelWriteProtokol(protokol);
        return "redirect:/";
    }
}
