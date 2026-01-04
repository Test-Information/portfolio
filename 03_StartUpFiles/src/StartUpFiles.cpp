// ==============================================================================================================
// <説明>
// 自宅環境でのC言語学習中に作成したexeファイルです。
// Ｃ言語のプロジェクトに参画開始したときに学習のために作成しました。
// 本プログラムの目的は、指定のアプリを起動させて、そのアプリのウィンドウを右、または左にスナップするものです。
// 
// 品質に問題がないことを確認したつもりではありますが、念のため実行前にコードチェックのうえ実行していただけますと幸いです。
//
// <作成情報>
// 作成者  ：神部 慶太
// 作成日  ：2023/02頃
// 開発期間：実用開始まで約７日、ポートフォリオ用のリファクタリング約１日
// 
// <使い方・および処理内容>
// 1)INIファイルに実行したいファイルを指定する
// 2)本プログラムをスタートアップで実行する
// 3)INIファイルに設定されているパスのアプリケーションが起動する
// 4)起動されたアプリケーションプロセスに「Win+←」、「Win+→」のキーを送信する。
// ※その後、(3)～(4)を繰り返し処理終了
// 
// 
// <発見バグ>
// 以下の対応はユーザーへ悪影響を与えないと考えエラーハンドリング処理を行わない
// ・アプリを起動してからユーザーが別のアプリをアクティブにして本アプリがスナップ（Win+矢印キー）を行ってしまった場合
// 
// <！！注意！！>
// 本プログラムはC言語による別プロセスの実行、およびキーイベントをOSに送るプログラムです。
// よってセキュリティソフトに悪意のあるプログラム扱いされる可能性があり、本プログラム実行時に削除されてしまう可能性がありますが、
// 決して悪意のあるプログラムではないことをあらかじめご了承ください。
// 採用ご担当者様の案件がC言語を扱う現場ではない場合、不要な技術ですので、実行しないことを推奨いたします。
// 万が一、起動する場合は、全てのプログラムを閉じた状態で実行することを推奨いたします。
// プログラムの途中で強制処理の処理を行っているが、その瞬間にユーザーがアクティブにしたプログラムを強制終了してしまう可能性があるため
// 
// バージョン：v.0.0.2
// ==============================================================================================================

#define WAIT_TIME1 3
#define WAIT_TIME2 1
#include <windows.h>
#include <stdio.h>

VOID keyExec(BYTE keyCode);                                 //プロトタイプ宣言
LONG windowSetup(TCHAR pName[], byte keyCode, PINT cnt);    //プロトタイプ宣言

//****************************************************************
//*【処理概要】   ：エントリーポイント
//* 第一パラメータ：
//* 第二パラメータ：
//* 第三パラメータ：コマンドライン引数
//* 第四パラメータ：
//* リターンコード：実行結果種別を返却
//*                   0:処理結果正常
//*                   1:処理結果異常 
//****************************************************************
int WINAPI WinMain(HINSTANCE hInst, HINSTANCE hPrevInst, LPSTR lpCmdLine, int intShowCmd) {
    TCHAR fName[300] = "\0";            // ファイル名
    TCHAR key[16] = "\0";               // iniファイル 取得対象キー
    TCHAR iniPath[MAX_PATH] = "\0";     // iniファイル パス
    TCHAR errBuff[256] = "\0";          // エラー情報格納用バッファ
    TCHAR ver[8] = "0.0.2\0";           // バージョン情報

    BYTE  keyCode;          // 物理キーボードのキーコード
    LONG  rCode = 0;        // 戻り値
    UINT  loopStart = 0;    // ループ開始
    UINT  loopEnd = 5;      // ループ終了
    UINT  hitFlag = 0;      // 実行プログラム存在フラグ

    // コマンドライン引数なし
    if (lpCmdLine == NULL || lpCmdLine[0] == '\0') {
        MessageBox(NULL, "INIファイルを引数に指定してください。\n"
            "半角スペースが含まれていないことを確認してください。", "エラー0", MB_OK | MB_ICONERROR);
        return 1;
    }
    else {
        // コマンドライン引数 最大バイト数取得
        size_t sz = strlen(lpCmdLine);
        // コマンドライン引数の値が260バイトを超える場合
        if (sz > MAX_PATH) {
            MessageBox(NULL, "扱い可能なパスの文字数を超えています。", "エラー1", MB_OK | MB_ICONERROR);
            return 1;
        }
        //コマンドライン引数の先頭がダブルクォーテーションの場合
        if (lpCmdLine[0] == 0x22/* " */) {
            //コマンドライン引数の末尾がダブルクォーテーションの場合    
            if (lpCmdLine[sz - 1] == 0x22/* " */) {
                // INIファイル名設定
                memcpy(iniPath, &lpCmdLine[1], sz - 2); // 先頭と末尾のダブルクォーテーションを引いた数を指定
                // 引数が２バイト以上の場合
                if (sz >= 2) {
                    iniPath[sz - 2] = '\0';             // ヌル止め設定(szが0以上の場合)
                }
            }
            else {
                // エラーメッセージ表示
                MessageBox(NULL, "コマンドライン引数をダブルクォーテーションで囲わないでください。", "エラー", MB_OK | MB_ICONERROR);
                return 1;
            }
        }
        else {
            // INIファイル名設定
            strcpy_s(iniPath, sizeof(iniPath), lpCmdLine);
        }
    }

    // ループ開始位置取得
    loopStart = GetPrivateProfileInt(
        "conf",                     // 対象セクション
        "loopStart",                // 対象キー
        0,                          // キー未存在時のデフォルト値
        iniPath);                   // iniファイルのパス

    // ループ終了位置取得
    loopEnd = GetPrivateProfileInt(
        "conf",                     // 対象セクション
        "loopEnd",                  // 対象キー
        5,                          // キー未存在時のデフォルト値
        iniPath);                   // iniファイルのパス

    // 上限値が２０を超えた場合
    if (loopEnd > 20) {
        loopEnd = 20;   // ループ最大値を２０に補正
    }

    // キー取得処理
    for (int cnt = loopStart; cnt < loopEnd; cnt++) {
        // キー名設定
        sprintf_s(&key[0], sizeof(key), "key%d", cnt);
        //INIファイル取得処理
        GetPrivateProfileString(
            "StartUpFile",              // 対象セクション
            key,                        // 対象キー
            "default",                  // キー未存在時のデフォルト値
            fName,                      // 取得結果の格納先
            sizeof(fName),              // 格納バッファサイズ
            iniPath);                   // iniファイルのパス

        // iniファイル キー取得結果がdefaultでない場合
        if (strcmp(fName, "default") != 0) {
            // キーの値が空だった場合
            if (strcmp(fName, "") == 0) {
                continue;           // 次ループ遷移
            }
            hitFlag += 1;
            // キー末尾が奇数の場合
            if ((cnt % 2) != 0) {
                keyCode = VK_LEFT;  // 左矢印キー
            }
            // キー末尾が偶数の場合
            else {
                keyCode = VK_RIGHT; // 右矢印キー
            }
            rCode = windowSetup(&fName[0], keyCode,&cnt);
            // 非アイドル状態
            if (rCode != 0) {
                return 1;
            }
        }
        else {
            // エラー情報格納
            sprintf_s(errBuff,
                sizeof(errBuff),
                "INIファイルから値を取得できませんでした。\n"
                "INIファイル格納場所が権限の影響を受けていないことを確認してください。\n"
                "取得対象:%s\n"
                "ループカウント:%d\n"
                "ver:%s",iniPath, cnt, ver);
            MessageBox(NULL, errBuff, "エラー2", MB_OK | MB_ICONERROR);
            return 1;   //取得結果が空だった場合終了
        }
    }
    // 実行プログラムなし
    if (hitFlag == 0) {
        // エラーメッセージ表示
        MessageBox(NULL, "実行プログラムがありませんでした。", "エラー3", MB_OK | MB_ICONERROR);
    }
    return 0;
}

//****************************************************************
//*【処理概要】   ：プロセスを起動させる
//* 第一パラメータ：起動ファイル名
//* 第二パラメータ：キーコード
//*                   VK_LEFT (0x25)：左矢印キーのキーコード
//*                   VK_RIGHT(0x27)：右矢印キーのキーコード
//* リターンコード：実行結果種別を返却
//*                   0:処理結果正常
//*                   1:処理結果異常 
//****************************************************************
LONG windowSetup(TCHAR pName[], byte keyCode,PINT cnt) {
    // CreateProcessパラメータ設定
    STARTUPINFO si;                         // ウィンドウ情報(サイズや位置など)
    PROCESS_INFORMATION pi;                 // 起動プロセスの識別情報

    // ウィンドウ設定初期化
    ZeroMemory(&si, sizeof(si));            // ウィンドウ情報 メモリクリア
    ZeroMemory(&pi, sizeof(pi));            // 識別情報 メモリクリア
    si.cb = sizeof(si);                     // 構造体サイズ
    si.dwFlags = STARTF_USESHOWWINDOW;      // メインウィンドウの外観設定
    si.wShowWindow = 1;                     // ウィンドウ追加情報

    // ファイルパス設定
    TCHAR fPath[MAX_PATH];                  // ローカル配列作成
    strcpy_s(fPath, sizeof(fPath), pName);  // ローカル配列コピー

    // エラーメッセージ
    TCHAR errBuff[256] = "\0";              //エラー情報格納用バッファ

    // プロセス起動成功時
    if (CreateProcess(NULL, fPath, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi)) {
        // プロセス起動後、アイドル状態になるまで待機
        DWORD  rtnCode = WaitForInputIdle(pi.hProcess, 10000);

        // 正常終了時
        if (rtnCode == 0) {
            // キーコード実行関数コール
            keyExec(keyCode);
        }
        else {
            // プロセス終了
            TerminateProcess(pi.hProcess, 1);   //  ほかのプロセスが直前で最前面になった場合、プロセスを終了させてしまう可能性がある。
            //プロセスハンドル解放
            CloseHandle(pi.hThread);
            CloseHandle(pi.hProcess);

            // エラー情報格納
            sprintf_s(errBuff,
                sizeof(errBuff),
                "以下のエラーの可能性があります。\n"
                "・起動アプリがGUIアプリでない\n"
                "・アプリ起動からタイムアウトが発生\n"
                "処理を中止します。\n"
                "ファイルパス：\n"
                "%s\n"
                "ループ数：%d\n",
                fPath,
                *cnt);
            //エラーメッセージ表示
            MessageBox(NULL, 
                errBuff,
                TEXT("エラー4"),
                MB_OK | MB_ICONINFORMATION);
            return 1;
        }
        //プロセスのハンドルを解放
        CloseHandle(pi.hThread);
        CloseHandle(pi.hProcess);
    }
    else {
        DWORD	dwError = GetLastError();   //エラーコード値取得
        LPVOID  lpMsgBuf;                   //エラーメッセージ格納用バッファ
        TCHAR   errBuff[256];               //エラー情報格納用バッファ

        FormatMessage(
            FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
            NULL,               // メッセージ定義
            dwError,            // エラーコード
            NULL,               // 言語ID
            (LPSTR)&lpMsgBuf,   // エラーメッセージ格納バッファ
            0,                  // バッファ挿入サイズ
            NULL);              // 

        // エラー情報格納
        sprintf_s(errBuff,
            sizeof(errBuff),
            "エラーコード:%d\n"
            "エラーメッセージ:%s\n"
            "指定ファイル:%s\n"
            "カウント:%d",
            dwError,
            (LPCTSTR)lpMsgBuf,
            fPath, *cnt);

        // エラー情報を使用していた場合
        if (lpMsgBuf) {
            // メモリ解放
            LocalFree(lpMsgBuf);
        }
        // エラー情報の表示
        MessageBox(NULL, errBuff, TEXT("エラー5"), MB_OK | MB_ICONINFORMATION);
        return(1);  // 異常終了
    }
    return(0);      // 正常終了
}

//****************************************************************
//*【処理概要】:ウィンドウ スナップ処理を行う
//* 第一パラメータ：キーコード
//*                   VK_LEFT (0x25)：左矢印キーのキーコード
//*                   VK_RIGHT(0x27)：右矢印キーのキーコード
//* リターンコード：なし
//* 課題：自動でキー押下時、アプリが最前面になっていないとキー(Win+矢印キー)の空振りが発生する
//****************************************************************
VOID keyExec(BYTE keyCode) {
    // アプリ起動待機処理
    Sleep(WAIT_TIME1 * 1000);                               // キーイベント前待機処理(指定数字×ミリ秒)

    //キー送信処理
    keybd_event(VK_LWIN, 0, 0, 0);                          // 左Ｗｉｎキーを押したままにする
    keybd_event(keyCode, 0, 0, 0);                          // 引数のキーを押したままにする
    keybd_event(keyCode, 0, KEYEVENTF_KEYUP, 0);            // 引数のキーを離す
    keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0);            // 左Ｗｉｎキーを離す

    // キー処理完了 待機処理
    Sleep(WAIT_TIME2 * 1000);                               // キーイベント後待機処理(指定数字×ミリ秒)
}
