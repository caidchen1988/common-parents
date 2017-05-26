package com.qtdbp.ui.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

/**
 * @author: caidchen
 * @create: 2017-05-26 9:04
 * To change this template use File | Settings | File Templates.
 */
@Controller
public class UIController {

    /**
     * 页面头部进度条效果
     * @return
     */
    @RequestMapping(value = "/ui/pace", method = RequestMethod.GET)
    public String pace() {
        return "ui/pace" ;
    }

    /**
     * 警告框
     * @return
     */
    @RequestMapping(value = "/ui/toastr", method = RequestMethod.GET)
    public String toastr() {
        return "ui/toastr" ;
    }
}
