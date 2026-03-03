package sample;

import org.apache.tomcat.jakartaee.commons.lang3.StringUtils;

import jakarta.servlet.http.HttpServletRequest;
/**
 * ユーザー入力フォーム格納DTO
 * 課題１：入力値がJavaScriptを使用していた場合、無効化する。
 * 課題２：E-Mail形式の判定を実装する。
 */
public class UserRegistForm {
	private String custName;
	private String email;
	private String password;
	private String rePassword;
	
	public UserRegistForm() {
		// デフォルトコンストラクタ
	}
	
	public UserRegistForm(HttpServletRequest request) {
	    this.custName = request.getParameter("cust-name");
	    this.email = request.getParameter("email");
	    this.password = request.getParameter("password");
	    this.rePassword = request.getParameter("re-password");
	}
	
	public String getCustName() {
		return custName;
	}
	public void setCustName(String custName) {
		this.custName = custName;
	}
	public String getEmail() {
		return email;
	}
	public void setEmail(String email) {
		this.email = email;
	}
	public String getPassword() {
		return password;
	}
	public void setPassword(String password) {
		this.password = password;
	}
	public String getRePassword() {
		return rePassword;
	}
	public void setRePassword(String rePassword) {
		this.rePassword = rePassword;
	}
	
	public void validate(StringBuilder sb) {
		/**
		 * 入力有無チェック
		 */
		if (StringUtils.isEmpty(this.getCustName())) {
			sb.append("顧客名が入力されていません。<br>");
		}
		
		if (StringUtils.isEmpty(this.getEmail())) {
			sb.append("メールアドレスが入力されていません。<br>");
		}
		
		if (StringUtils.isEmpty(this.getPassword())) {
			sb.append("パスワードが入力されていません。<br>");
		}
		else if (!StringUtils.isAlphanumeric(this.getPassword())) {
			sb.append("パスワードは英数字のみで入力してください。<br>");
		}
		
		if (StringUtils.isEmpty(this.getRePassword())) {
			sb.append("パスワード（確認用）が入力されていません。<br>");
		}
		else if (!StringUtils.isAlphanumeric(this.getRePassword())) {
			sb.append("パスワード（確認用）は英数字のみで入力してください。<br>");
		}
		
		/**
		 * パスワードとパスワード（確認用）比較
		 */
		if (!(StringUtils.equals(this.getPassword(), this.getRePassword()))) {
			sb.append("パスワードが一致しません。<br>");
		}
		
		
	}
}
	