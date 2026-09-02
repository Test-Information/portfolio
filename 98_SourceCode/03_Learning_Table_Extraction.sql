--別紙 スキル詳細シートの画面証跡シートで学習ログを可視化するために作成したOracle SQL文

--#１．付与オブジェクト権限抽出クエリー#
SELECT
    ROWNUM AS "No.",
    NVL(au.username,'該当ユーザーなし') AS "自作ユーザー",
    NVL(ap.table_name,'未設定') AS "オブジェクト名",
    NVL(ap.grantor,'未設定') AS "権限付与実行ユーザ",
    NVL(ap.privilege,'未設定') AS "オブジェクト権限",
    NVL(ap.grantable,'未設定') AS "GRANT OPTION付与",
    NVL(TO_CHAR(au.created,'YYYY"年"MM"月"DD"日" HH24"時"MI"分"'),'未設定') AS "ユーザー作成日"
FROM all_users au
LEFT JOIN all_tab_privs ap
ON  au.username = ap.grantee
WHERE au.created BETWEEN TO_DATE('2025-10-01','YYYY-MM-DD') AND TO_DATE('2026-03-01','YYYY-MM-DD') --ユーザー作成期間
ORDER BY "No.";

**--       No. 自作ユーザー オブジェクト名  権限付与実行ユーザ   オブジェクト権限     GRANT OPTION付与     ユーザー作成日
---------- ------------ --------------- -------------------- -------------------- -------------------- ---------------
**--         1 PDBADMIN     未設定          未設定               未設定               未設定               2025-11-01
**--         2 TEST         未設定          未設定               未設定               未設定               2025-11-24
**--         3 TTEST        未設定          未設定               未設定               未設定               2025-11-17
**--         4 TEST_SCM     EMPLOYEE        SYS                  ALTER                NO                   2025-11-20
**--         5 TEST_SCM     EMPLOYEE        SYS                  DELETE               NO                   2025-11-20
**--         6 TEST_SCM     EMPLOYEE        SYS                  INSERT               NO                   2025-11-20
**--         7 TEST_SCM     EMPLOYEE        SYS                  SELECT               NO                   2025-11-20
**--         8 TEST_SCM     EMPLOYEE        SYS                  UPDATE               NO                   2025-11-20
**--         9 TEST_SCM     EXT_DATA        SYS                  EXECUTE              YES                  2025-11-20
**--        10 TEST_SCM     EXT_DATA        SYS                  READ                 YES                  2025-11-20
**--        11 TEST_SCM     EXT_DATA        SYS                  WRITE                YES                  2025-11-20
**--        12 SCM          未設定          未設定               未設定               未設定               2025-11-20
**--        13 USER1        未設定          未設定               未設定               未設定               2025-12-10
**--        14 TEST02       未設定          未設定               未設定               未設定               2025-11-24
**--        15 USER2        未設定          未設定               未設定               未設定               2025-12-10

**--15行が選択されました。
**--    経過: 00:00:01.00


--#２．学習テーブル抽出クエリー#
SELECT
    ROWNUM AS "No.",
    at.*
FROM
    (SELECT owner || 'で作成したテーブル数：' AS "集計対象", COUNT(*) AS "集計数"
    FROM all_tables --全テーブル格納テーブル
    WHERE tablespace_name = 'USERS' --表領域指定
    GROUP BY owner) at;
    

**--       No. 集計対象                             集計数
**------------ ------------------------------ ------------
**--         1 TEST_SCMで作成したテーブル数：           35
**--         2 TEST02で作成したテーブル数：              1
**--         3 USER1で作成したテーブル数：               1
**--
**--経過: 00:00:00.27
**--SQL>