package com.qtdbp.base.controller;

import com.qtdbp.base.utils.RequestMessage;
import com.qtdbp.base.utils.ResponseMessage;
import org.springframework.messaging.handler.annotation.MessageMapping;
import org.springframework.messaging.handler.annotation.SendTo;
import org.springframework.messaging.simp.SimpMessagingTemplate;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import javax.annotation.Resource;

/**
 * websocket Controller
 *
 * @author: caidchen
 * @create: 2017-05-31 10:01
 * To change this template use File | Settings | File Templates.
 */
@Controller
public class GreetingController {

    @RequestMapping(value = "/ws", method = RequestMethod.GET)
    public String index(){
        return "index";
    }

    /**
     * @MessageMapping定义一个消息的基本请求，功能也跟@RequestMapping类似
     * @param message
     */
    @MessageMapping("/welcome")
    @SendTo("/topic/notice")
    public ResponseMessage say(RequestMessage message){

        System.out.println(message.getName());
        return new ResponseMessage("welcome," + message.getName() + " !");
    }
}
