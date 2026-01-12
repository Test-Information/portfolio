package com.example.demo;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
/**
 * メニュー画面コントローラー
 * メニュー画面の動作を制御します。
 */
@Controller
public class MainController {
	/** メニュー画面を表示します */
	@GetMapping("/")
    public String index() {
        return "index";
    }
}
