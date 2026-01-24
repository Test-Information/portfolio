package com.example.demo;

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
 *  2)0円の場合
 *  4)テキストボックスがデベロッパーツールで改変された場合)
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
    		@RequestParam(required = false) String name,
    		@RequestParam(required = false) Long balance,
    		Model model) {

    	Customer cust =	 new Customer();	// 顧客情報作成モデルの生成
    	String errMsg = "";								// エラーメッセージ初期化
    	char[] invalidChars = { '!', '"', '#', '$', '%', '&', '\'', '(', ')', '=', '~', '|', '`', '{', '+', '*', '}', '<', '>', '?', '_' 	};	// 入力値として無効な文字一覧
    	
    	/** 顧客名 入力値有無 判定 */
    	if(!(StringUtils.hasText(name))) {
    		errMsg += "顧客名が入力されていません\n";
    	} else {
        	outerLoop:
        	for(int cnt = 0; cnt < name.length(); cnt++) {
            	for(int cnt2 = 0 ; cnt2 < invalidChars.length;cnt2++) {
    				if(name.charAt(cnt) == invalidChars[cnt2]) {
    					errMsg += "記号が入力されています。【" + name.charAt(cnt) + "】\n";
    					break outerLoop;
    				}
            	}
        	}
    	}
    	
    	/** 残高 入力値有無 判定 */
    	if(balance  == null) {
    		errMsg += "金額が入力されていません\n";
    	}else if(balance == 0) {
    		errMsg += "金額が０円です\n";
    	}else if(balance < 0) {
    		errMsg += "金額が負の入力値です\n";
    	}
    	
    	/** 各フィールド 入力値設定済 */
    	if(errMsg.isEmpty()){
    		cust.setCustName(name);			// 顧客名設定
    		cust.setCustBalance(balance);	// 残高設定
            repository.save(cust);					// エンティティ更新
            model.addAttribute("result", "顧客情報の作成に成功しました");
            return "customer";
    	} else {
    		model.addAttribute("result", errMsg);	//各エラーメッセージ 一元設定
    	}
        return "customer";
    }
}