package sample;

import java.io.IOException;

import jakarta.servlet.RequestDispatcher;
import jakarta.servlet.ServletException;
import jakarta.servlet.annotation.WebServlet;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;

/**
 * ユーザー登録処理
 * 
 */
@WebServlet("/CustomerRegistration")
public class CustomerRegistration extends HttpServlet {
	private static final long serialVersionUID = 1L;

	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        RequestDispatcher dispatchr = request.getRequestDispatcher("index.jsp");
		dispatchr.forward(request, response);
	}
	
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {

		String fwdPage;
		String[] invalidChars = { "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "-", "_", "=", "+", "[", "]", "{", "}"};
		
		// エラーメッセージ格納処理
		StringBuilder sb = new StringBuilder();
		
		// ユーザー登録フォーム入力値格納＆バリデーション
		UserRegistForm urf = new UserRegistForm(request);
		urf.validate(sb);
		
		// エラーメッセージありの場合
		if(sb.length() > 0) {
			// 前後タグ設定
			sb.insert(0,"<p class=\"err-msg\">").append("</p>");

			// 画面表示値 設定
			request.setAttribute("cust-name", urf.getCustName());
			request.setAttribute("emal", urf.getEmail());
			request.setAttribute("errorMsg",sb.toString());
			
			// 遷移先ページ
			fwdPage = "index.jsp";
		}
		else {
			// 遷移先ページ
			fwdPage = "registSuccess.jsp";
		}
		
		// フォワード処理
		RequestDispatcher dispatchr = request.getRequestDispatcher(fwdPage);
		dispatchr.forward(request, response);

	}

}
