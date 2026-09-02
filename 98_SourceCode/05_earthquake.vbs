' ==============================================================================================================
'  プログラム名：【地震速報表示 スクリプト】
'  【説明】
'  自宅環境で使用しているスクリプトです。
'  実行されたときに地震のリアル地震速報、ニュース、動画サイト、津波のニュースに関するサイトをブラウザで表示します。
'  品質に問題がないことを確認したつもりではありますが、念のため実行前にコードチェックのうえ実行していただけますと幸いです。
'
'  【作成情報】
'  作成者  ：神部 慶太
'  作成日  ：2022/11/20
'  
'  【使い方、および実行処理】
'  1)ブラウザ変数に指定のブラウザのexeファイルのパスを設定（chromeまたはedge）
'  2)本プログラムのショートカットファイルを作成し、プロパティで呼び出し用のショートカットキーを設定
'  3)ショートカットファイルをスタートメニュー配下にデプロイする
'  4)ショートカットキーなどで呼び出すと設定された地震に関するWebページを表示する
'  
'  【試験】
'  brwsの値が存在する場合と、存在しない場合の処理のみ確認済
' ==============================================================================================================

'変数宣言
Dim objCR
'ブラウザ
Dim brws : brws = "C:\Program Files\Google\Chrome\Application\chrome.exe"
'リアルタイム震度
Dim url1 : url1 = "https://typhoon.yahoo.co.jp/weather/jp/earthquake/kyoshin/"
'Yahoo天気・震害（アンカーリンク）
Dim url2 : url2 = "https://typhoon.yahoo.co.jp/weather/jp/earthquake/#cat-pass"
'Youtube 検索ワード：地震
Dim url3 : url3 = "https://www.youtube.com/results?search_query=%E5%9C%B0%E9%9C%87&sp=EgIIAg%253D%253D"
'Google 通常検索ページ 検索ワード：地震 ※最新１時間のみ表示
Dim url4 : url4 = "https://www.google.com/search?q=%E5%9C%B0%E9%9C%87&source=lnt&tbs=qdr:h&sa=X&ved=2ahUKEwjy5rf43c_2AhWSNKYKHY16BG0QpwV6BAgBEBo&biw=1920&bih=929&dpr=1"
'Google NEWSページ     検索ワード：地震 ※最新１時間のみ表示
Dim url5 : url5 = "https://www.google.com/search?q=%E5%9C%B0%E9%9C%87&tbm=nws&source=lnt&tbs=qdr:h&sa=X&ved=2ahUKEwjE7rf43c_2AhWRvZQKHcFBC3AQpwV6BAgBEBo&biw=1920&bih=929&dpr=1"
'Google NEWSページ     検索ワード：津波 ※最新１時間のみ表示
Dim url6 : url6 = "https://www.google.com/search?q=%E6%B4%A5%E6%B3%A2&tbm=nws&source=lnt&tbs=qdr:h&sa=X&ved=2ahUKEwjwjPiX38_2AhU2w4sBHZVLDa0QpwV6BAgBEBo&biw=1920&bih=929&dpr=1"

'WshShellオブジェクトのオブジェクトを作成する
Set objCR = WScript.CreateObject ("WSCript.shell")
'chrome.exeを実行し、URLを引数で渡す
Dim cmd : cmd = """" & brws & """ --new-window " & url1 & " " & url2 & " " & url3 & " " & url4 & " " & url5 & " " & url6
' 実行
objCR.Run cmd

'変数解放（記述しなくても問題なし）
Set objCR = Nothing
