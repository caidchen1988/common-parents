package com.qtdbp.ui.controller;

import org.springframework.stereotype.Controller;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

/**
 * 表单
 *
 * @author: caidchen
 * @create: 2017-05-25 14:21
 * To change this template use File | Settings | File Templates.
 */
@Controller
public class FormController {

    @RequestMapping(value = "/form/validate", method = RequestMethod.GET)
    public String getValidate() {
        return "form/validate" ;
    }


    @ResponseBody
    @RequestMapping(value = "/form/validate", method = RequestMethod.POST)
    public ModelMap postValidate(HttpServletRequest request) {

//        ModelAndView view = new ModelAndView("dashboard") ;

        ModelMap map = new ModelMap() ;

        Map<String, String[]> result = request.getParameterMap() ;
        if(result != null) {
            for (String key : result.keySet()) {
//                view.addObject(key, result.get(key));

                map.put(key, result.get(key)) ;
System.out.println("key: "+ key +", Object: "+ result.get(key));
            }
        }

        return map ;
    }
}
