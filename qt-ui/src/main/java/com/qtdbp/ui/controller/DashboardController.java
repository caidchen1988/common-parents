package com.qtdbp.ui.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

/**
 * @author: caidchen
 * @create: 2017-05-25 16:39
 * To change this template use File | Settings | File Templates.
 */
@Controller
public class DashboardController {

    @RequestMapping(value = "/", method = RequestMethod.GET)
    public String getValidate() {
        return "dashboard" ;
    }

}
