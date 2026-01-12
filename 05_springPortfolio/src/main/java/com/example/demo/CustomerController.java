package com.example.demo;

import java.util.Optional;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;


/**
 * 顧客情報作成画面コントローラー
 * 
 * 顧客情報作成画面への初回画面遷移、および顧客情報作成画面からの送信データ受信時の処理を制御します。
 * 
 * 【今後の修正予定】
 * ポートフォリオとして公開を優先するため以下は公開後の対応を行う予定とする
 * １．バリデーション処理の追加
 *  1)記号あり
 *  2)0円の場合
 *  3)ゼロパディングされた金額の場合
 *  4)テキストボックスがデベロッパーツールで改変された場合
 */
@Controller
public class CustomerController {
    private final CustomerRepository repository;

    public CustomerController(CustomerRepository repository) {
        this.repository = repository;
    }
    
    /**
     * Getメソッド受信時の処理
     * 初回画面表示時の処理を行います。
     * @return 顧客作成画面ビュー
     */
    @GetMapping("/customer")
    public String getCustomer() {
    	return "customer";
    }
    /**
     * Postメソッド受信時の処理
     * 顧客作成画面から顧客作成ボタンが押されたときの処理を行います。
     * @param name		顧客名フィールドの入力値
     * @param balance		残高フィールドの入力値
     * @param model		モデル
     * @return					顧客作成画面ビュー
     */
    @PostMapping("/customer")
    public String postCustomer(
    		@RequestParam (required = false) String name,
    		@RequestParam(required = false) Long balance,
    		Model model) {
    	
    	Customer cust =	 new Customer();	// 顧客情報作成モデルの生成
    	String errMsg = "";								// エラーメッセージ初期化

    	Optional<String> optName = Optional.ofNullable(name);
    	Optional<Long> optBalance = Optional.ofNullable(balance);

    	// 顧客名 入力値有無の判定
    	if(optName.isPresent()) {
    		cust.setName(name);	
    	} else {
    		errMsg += "顧客名が入力されていません\n";
    	}
    	
    	// 残高 入力値有無の判定
    	if(optBalance.isPresent())  {
    		cust.setBalance(balance);
    	} else {
    		errMsg += "金額が入力されていません\n";
    	}
    	
    	// 各フィールドの入力値が全て設定されていた場合
    	if((StringUtils.hasText(cust.getName())) && (cust.getBalance() != null)){
            repository.save(cust);
            model.addAttribute("result", "顧客情報の作成に成功しました");
            return "customer";
    	} else {
    		model.addAttribute("result", errMsg);	//各エラーメッセージの一元設定
    	}
        return "customer";
    }
}