<!-- デザインに関してはインターネットのものを流用していますので、評価はJavaだけしていただければと思います。 -->>
<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>ユーザー登録画面</title>
<link rel="stylesheet" type="text/css" href="${pageContext.request.contextPath}/css/style.css">
</head>
<body>
<div class="contact-form">
  <div class="form-header">	
    <h2>ユーザー登録画面</h2>
    <p>ユーザー情報を入力してください。</p>
  </div>
  <form id="contactForm" action="CustomerRegistration" method="POST">
    <c:if test="${not empty errorMsg}">
      <div>${errorMsg}</div>
    </c:if>
    
    <div class="form-group">
      <label for="cust-name" class="form-label">
        お名前 <span class="required">*</span>
      </label>
      <input type="text" class="form-control" id="cust-name" name="cust-name" maxlength="50" value="${custName}">
    </div>

    <div class="form-group">
      <label for="email" class="form-label">
        メールアドレス <span class="required">*</span>
      </label>
      <!--<input type="email" class="form-control" id="email" name="email" required maxlength="100">-->
      <input class="form-control" id="email" name="email" maxlength="100">
    </div>
    
    <div class="form-group">
      <label for="password" class="form-label">
        パスワード <span class="required">*</span>
      </label>
      <input type="password" class="form-control" id="password" name="password" maxlength="100">
    </div>
    
    <div class="form-group">
      <label for="re-password" class="form-label">
        パスワード（確認用） <span class="required">*</span>
      </label>
      <input type="password" class="form-control" id="re-password" name="re-password" maxlength="100">
    </div>
    <button type="submit" class="btn-submit">
      送信する
    </button>
  </form>
</div>
</body>
</html>