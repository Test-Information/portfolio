package com.example.demo;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
/**
 * アプリケーション起動クラス
 */
@SpringBootApplication
public class SprApplication {
	/**
	 * メインメソッド
	 * @param args コマンドライン引数
	 */
	public static void main(String[] args) {
		SpringApplication.run(SprApplication.class, args);
	}
}
