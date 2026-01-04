Attribute VB_Name = "Module1"
'' ==============================================================================================================
' <説明>
' 自宅環境で月末書類の勤務表や報告書を自動で初期化して、名前を付けて保存するExcel VBAです。
' Excel VBAの学習のために作成しました。
' 品質に問題がないことを確認したつもりではありますが、念のため実行前にコードチェックのうえ実行していただけますと幸いです。
'
' <作成情報>
' 作成者  ：神部 慶太
' 作成日  ：2022/04頃
' 開発期間：実用開始まで約4日、ポートフォリオ用のリファクタリング約１日
'
' <使い方・および処理内容>
' 1)マスタシートで月を設定
' 2)必要に応じてパラメータ設定
' 3)勤怠初期化ボタン押下
' 4)勤務表シート(YYYY年MM月)のA/B列（日付）を自動で入力
' 5)勤務表シート(YYYY年MM月)のC/D列（出社・退社時間）を設定した値で入力
' 6)本ブックの格納ディレクトリに「月末書類」フォルダを作成
' 7)「月末書類」フォルダ直下にパラメータで設定された「YYYYMM」フォルダを作成
' 8)「YYYYMM」フォルダに勤務表シートを新規ブックで保存
' 9)「YYYYMM」フォルダにダミー報告書シートを新規ブックで保存
'
'
' <！！注意！！>
' 1)他ブックに影響を与えないようにソースコードをチェックしたつもりではありますが、
' より安全に実行していただくためには、他ブックを全て閉じておくことを推奨いたします。
' 2)本ブックと同じディレクトリにフォルダを作成します。たまたま同名フォルダ、さらに同名ファイルまで存在していた場合、上書きしてしまう可能性があります。
' 本ブックを格納するディレクトリは配下にフォルダが存在しないディレクトリが推奨されます。
' 3)セキュリティ上の観点からダウンロードしたxlsmのマクロ実行が気になる場合は動画をご覧ください。
'
' バージョン：v.0.0.1
' ==============================================================================================================

Sub 勤怠リセット()
    '**********************************************************************
    '変数宣言
    '**********************************************************************
    Dim month As Date        '月
    Dim sheetsName As String 'シート名
    Dim startTime As Date    '出社時間
    Dim endTime As Date      '退社時間
    Dim lastDay As Date      '月末最終日
    Dim breakTime As Date    '休憩時間
    Dim closeFlag As Boolean 'ブックを閉じるフラグ
    Dim openFlag As Boolean  'ブックを開いたままにするかのフラグ
    Dim folder As String     'フォルダ名
    Dim saveWB As Workbook   '保存するワークブック
    
    '**********************************************************************
    '変数代入
    '**********************************************************************
    month = ThisWorkbook.Worksheets("マスタ").Range("d9").Value       'マスタシートから値を格納
    startTime = ThisWorkbook.Worksheets("マスタ").Range("d10").Value  'マスタシートから値を格納
    endTime = ThisWorkbook.Worksheets("マスタ").Range("d11").Value    'マスタシートから値を格納
    sheetsName = ThisWorkbook.Worksheets("マスタ").Range("d12").Value 'マスタシートから値を格納
    breakTime = ThisWorkbook.Worksheets("マスタ").Range("d13").Value  'マスタシートから値を格納
    lastDay = ThisWorkbook.Worksheets("マスタ").Range("d14").Value    'マスタシートから値を格納
    closeFlag = ThisWorkbook.Worksheets("マスタ").Range("d15").Value  'マスタシートから値を格納
    openFlag = ThisWorkbook.Worksheets("マスタ").Range("d16").Value   'マスタシートから値を格納
    folder = ThisWorkbook.Worksheets("マスタ").Range("d17").Value     'マスタシートから値を格納
    ThisWorkbook.Worksheets(2).Name = sheetsName                      '勤務表(YYYY年MM月)シート　シート名取得
    
    
    '**********************************************************************
    '勤怠シート セル値クリア処理
    '**********************************************************************
    With ThisWorkbook.Worksheets(sheetsName)
        '勤務表(YYYY年MM月)シート　初期化処理
        .Range("a5:a35").ClearContents '日付列(A5～A35)削除
        .Range("A4").Value = month     '年月セルに値をセット(A4)
        .Range("c5:c35") = startTime   '出社時間セット(C5～C35)
        .Range("d5:d35") = endTime     '退社時間セット(D5～D35)
        .Range("g5:g35") = breakTime   '休憩時間セット(G5～G35)
    End With
            
    '**********************************************************************
    '勤怠シート 日付列 初期化処理
    '**********************************************************************
    '勤怠シート　A5セルから存在日付数分の日付をセット
    For i = 0 To Day(lastDay) - 1
        ThisWorkbook.Worksheets(sheetsName).Range("A" & i + 5).Value = month + i
    Next
    
    '**********************************************************************
    '勤怠シート 曜日列 初期化処理
    '**********************************************************************
    '勤怠シート　B列　オートフィル
    ThisWorkbook.Worksheets(sheetsName).Activate '勤務表(YYYY年MM月)シート アクティブ化
    ThisWorkbook.Worksheets(sheetsName).Range("B5").Select 'B5セル 選択状態
    Selection.AutoFill Destination:=ThisWorkbook.Worksheets(sheetsName).Range("B5:B35"), Type:=xlFillDefault
    
    '**********************************************************************
    '勤怠シート 曜日列 土日処理
    '**********************************************************************
    With ThisWorkbook.Worksheets(sheetsName)
        For cnt = 5 To 35
            If .Range("B" & cnt) = "土" Or .Range("b" & cnt) = "日" Then
                'B列の値が土日だった場合
                .Range("C" & cnt).ClearContents 'C列該当セル値クリア
                .Range("D" & cnt).ClearContents 'C列該当セル値クリア
                .Range("G" & cnt).ClearContents 'C列該当セル値クリア
            End If
        
            If .Range("A" & cnt) = "" Then
                'A列の値が空だった場合（29日～31日）
                .Range("B" & cnt).ClearContents 'B列該当セル値クリア
                .Range("C" & cnt).ClearContents 'C列該当セル値クリア
                .Range("D" & cnt).ClearContents 'D列該当セル値クリア
                .Range("G" & cnt).ClearContents 'G列該当セル値クリア
            End If
        Next cnt
    End With
    
    '**********************************************************************
    '月次報告書シート
    '**********************************************************************
    '月次報告書 記入日
    With ThisWorkbook.Worksheets("ダミー報告書")
        .Range("c4").Value = lastDay '記入日セルへ月末最終日をセット
        .Activate          'アクティブ状態に変更
        Range("c4").Select 'C4セルにフォーカス位置セット
    End With
    
    '**********************************************************************
    '* デフォルトフォーカス位置セット
    '**********************************************************************
    Application.Goto ThisWorkbook.Worksheets(sheetsName).Range("D5") '勤怠シート   D5セルにフォーカス位置セット
    Application.Goto ThisWorkbook.Worksheets("マスタ").Range("B2")   'マスタシート B2セルにフォーカス位置セット

    'バックアップ用フォルダ作成
    If Dir(ThisWorkbook.Path & "\月末書類\" & folder, vbDirectory) = "" Then
        '本ブック格納ディレクトリ直下の月末書類フォルダ内に指定YYYYMMフォルダが存在しない場合
        MkDir ThisWorkbook.Path & "\月末書類\" & folder 'YYYYMMフォルダ作成
    End If
    
    ' 保存時の警告（上書き確認）を無効化
    Application.DisplayAlerts = False

    '勤怠シート取得
    ThisWorkbook.Worksheets(sheetsName).Copy '勤怠シート状態コピー
    Set saveWB = ActiveWorkbook 'コピーしたシートを変数に格納
    MsgBox "保存先:" & vbCrLf & ThisWorkbook.Path & "\月末書類\" & folder & "\勤務表.xlsx"

    'YYYYMMフォルダに新規ブックとして保存
    saveWB.SaveAs Filename:=ThisWorkbook.Path & "\月末書類\" & folder & "\勤務表.xlsx"
    
    MsgBox "勤務表シートを新規保存完了"
    
    'ブック開閉処理
    If openFlag = False Then
        saveWB.Close SaveChanges:=False '保存せず閉じる
    End If
    
    'ダミー報告書シート取得
    ThisWorkbook.Worksheets("ダミー報告書").Copy
    Set saveWB = ActiveWorkbook
    saveWB.SaveAs Filename:=ThisWorkbook.Path & "\月末書類\" & folder & "\ダミー報告書（月報）.xlsx"
    MsgBox "ダミー報告書（月報）を新規保存完了。"
    
    '保存時の警告を（上書き確認）有効化
    Application.DisplayAlerts = True
    
    'ブック開閉処理
    If openFlag = False Then
        saveWB.Close False
    End If

    MsgBox "本ブックを終了します"
    
    'マスタ開閉処理
    If closeFlag = True Then
        ThisWorkbook.Close SaveChanges:=True '本ブックを保存して閉じる
    End If
    
End Sub

