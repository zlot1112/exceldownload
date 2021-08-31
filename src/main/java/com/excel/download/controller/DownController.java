package com.excel.download.controller;

import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.servlet.ModelAndView;

@Slf4j
@Controller
public class DownController {

    @PostMapping("/down.do")
    public ModelAndView down() {
        ModelAndView mav = new ModelAndView();
        mav.setViewName("ExcelDownView");
        return mav;
    }

}
