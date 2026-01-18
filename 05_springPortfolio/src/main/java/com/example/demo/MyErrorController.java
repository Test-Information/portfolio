package com.example.demo;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import org.springframework.boot.web.servlet.error.ErrorController;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import jakarta.servlet.RequestDispatcher;
import jakarta.servlet.http.HttpServletRequest;

@Controller
@RequestMapping("/error")
public class MyErrorController implements ErrorController {
	final ErrLogRepository repository;
	
	public MyErrorController(ErrLogRepository repository) {
		this.repository = repository;
	}
	
	/**
	   * HTML レスポンス用の ModelAndView オブジェクトを返す。
	   *
	   * @param req リクエスト情報
	   * @param mav レスポンス情報
	   * @return HTML レスポンス用の ModelAndView オブジェクト
	   */
	@RequestMapping(produces = MediaType.TEXT_HTML_VALUE)
	public ModelAndView myErrorHtml(HttpServletRequest req, ModelAndView mav) {
		//日付設定
		LocalDateTime  curDate = LocalDateTime .now();																//	現在日付取得 
		DateTimeFormatter fmt = DateTimeFormatter.ofPattern("yyyy年MM月dd日 HH:mm:ss");	//	フォーマット宣言

		// HTTP ステータス
		HttpStatus status = getHttpStatus(req);
 
		// モデル設定
		mav.setStatus(status);			// HTTP ステータス セット
		mav.setViewName("error");	// ビュー名 error.html 
		mav.addObject("eDate" , curDate.format(fmt));	//	エラー日付
		mav.addObject("status" , status.value());				// ステータス設定
		mav.addObject("path"   , req.getAttribute(RequestDispatcher.ERROR_REQUEST_URI));	// 発生URI
		
		// エラーログ テーブル設定
		ErrLog eLog = new ErrLog();
		eLog.setEDate(curDate);
		eLog.setStatus(status.value());
		eLog.setPath(req.getAttribute(RequestDispatcher.ERROR_REQUEST_URI).toString());
		repository.save(eLog);
 
		// ステータス判定処理
		switch(status) {
			/* 404 */
			case NOT_FOUND:
				mav.addObject("message", "ページが見つかりません。URIをご確認ください。"); 
				break;
			default :
				mav.addObject("message", "システムエラーが発生しました。システム管理者にお問い合わせ下さい。");
		}
		return mav;
	}
	/**
	 * HTTPステータス 判定処理
	 * 
	 * @param req    リクエスト情報
	 * @return 処理成功時：該当HTTP ステータスコード
	 *                処理失敗時：サーバーエラー ステータスコード
	 */
	private static HttpStatus getHttpStatus(HttpServletRequest req) {
		// HTTP ステータス設定
		Object statusCode = req.getAttribute(RequestDispatcher.ERROR_STATUS_CODE);
		// 型・存在チェック
		if(statusCode instanceof Integer) {
			HttpStatus status = HttpStatus.resolve((int)statusCode);
			 if(status != null) {
					return status;
			 }
		}
		// エラーチェック不正時 デフォルト値
		return HttpStatus.INTERNAL_SERVER_ERROR;
	}
}
		
