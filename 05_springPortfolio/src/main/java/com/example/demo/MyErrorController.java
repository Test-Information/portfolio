package com.example.demo;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.springframework.boot.web.error.ErrorAttributeOptions;
import org.springframework.boot.web.servlet.error.DefaultErrorAttributes;
import org.springframework.boot.web.servlet.error.ErrorController;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.context.request.ServletWebRequest;
import org.springframework.web.servlet.ModelAndView;

import jakarta.servlet.RequestDispatcher;
import jakarta.servlet.http.HttpServletRequest;

@Controller
@RequestMapping("/error")
public class MyErrorController implements ErrorController {
 
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
		LocalDate curDate = LocalDate.now();																//	現在日付取得 
		DateTimeFormatter fmt = DateTimeFormatter.ofPattern("yyyy年MM月dd日");	//	フォーマット宣言

		// HTTP ステータス
		HttpStatus status = getHttpStatus(req);
 
		// モデル設定
		mav.setStatus(status);			// HTTP ステータス セット
		mav.setViewName("error");	// ビュー名 error.html 
		mav.addObject("eDate" , curDate.format(fmt));	//	エラー日付
		mav.addObject("status" , status.value());				// ステータス設定
		mav.addObject("path"   , req.getAttribute(RequestDispatcher.ERROR_REQUEST_URI));	// 発生URI
 
		// ステータス判定処理
		switch(status.value()) {
			case 404:
				mav.addObject("message", "ページが見つかりません。"); 
				break;
			default :
				mav.addObject("message", "システムエラーが発生しました。システム管理者にお問い合わせ下さい。");
		}
		return mav;
	}
 
	/**
	 * JSON レスポンス用の ResponseEntity オブジェクトを返す。
	 *
	 * @param req リクエスト情報
	 * @return JSON レスポンス用の ResponseEntity オブジェクト
	 */
	@RequestMapping
	public ResponseEntity<Map<String, Object>> myErrorJson(HttpServletRequest req) {
 
		// エラー情報を取得
		Map<String, Object> attr = getErrorAttributes(req);
 
		// HTTP ステータスを決める
		HttpStatus status = getHttpStatus(req);
 
		// 出力したい情報をセットする
		Map<String, Object> body = new HashMap();
		body.put("status", status.value());
		body.put("timestamp", attr.get("timestamp"));
		body.put("error", attr.get("error"));
		body.put("exception", attr.get("exception"));
		body.put("message", attr.get("message"));
		body.put("errors", attr.get("errors"));
		body.put("trace", attr.get("trace"));
		body.put("path", attr.get("path"));
 
		// 情報を JSON で出力する
		return new ResponseEntity<>(body, status);
	}
 
	/**
	  * JSON レスポンス用の エラー情報を抽出する。
	  *
	  * @param req リクエスト情報
	  * @return エラー情報
	  */
	private Map<String, Object> getErrorAttributes(HttpServletRequest req) {
		// DefaultErrorAttributes クラスで詳細なエラー情報を取得する
		ServletWebRequest swr = new ServletWebRequest(req);
		DefaultErrorAttributes dea = new DefaultErrorAttributes();
		ErrorAttributeOptions eao = ErrorAttributeOptions.of(
				ErrorAttributeOptions.Include.BINDING_ERRORS,
				ErrorAttributeOptions.Include.EXCEPTION,
				ErrorAttributeOptions.Include.MESSAGE,
				ErrorAttributeOptions.Include.STACK_TRACE);
		return dea.getErrorAttributes(swr, eao);
	}
 
	/**
	 * レスポンス用の HTTP ステータスを決める。
	 *
	 * @param req リクエスト情報
	 * @return レスポンス用 HTTP ステータス
	 */
	private static HttpStatus getHttpStatus(HttpServletRequest req) {
		// HTTP ステータスを決める
		// ここでは 404 以外は全部 500 にする
		Object statusCode = req.getAttribute(RequestDispatcher.ERROR_STATUS_CODE);
		HttpStatus status = HttpStatus.INTERNAL_SERVER_ERROR;
		if (statusCode != null && statusCode.toString().equals("404")) {
			status = HttpStatus.NOT_FOUND;
		}
		return status;
	}
 
}
		
