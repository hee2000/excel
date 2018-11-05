/** 作業報告書出力
 *  ユーザ定義関数を利用してExcel印刷のレポートを出力するサンプルです。
 *　１．ユーザIDとユーザ名で改シートします
 *　２．週単位に作業時間を小計欄に出力します
 *　３．月単位の作業時間を合計欄に出力します
 *  印刷の手順は
 *　１．帳票テンプレート画面にExcelテンプレートをアップします。
 *　２．画面設定でボタンを追加して保存します（印刷ボタンではありません）
 *　３．ボタンに新規関数を追加して保存します。
 *　４．下記の関数内容を追加して保存します。
 *　５．ユーザ確認環境で、印刷処理結果を確認します。
 */

//F20160622103106831
var excelDataBean = null;
var sheeBean = null;
var excelPrint = getExcelPrint();
var strSql = new StringBuilder();
var al = null;
var row = null;
var row1 = null;
var rowDataList = null;
// 社員ID
var userId = "";
// 社員名
var userNm = "";
//勤務年月
var sYearMonth = getObj('startDate'); 
//var eYearMonth = getObj('endDate'); 
//チーム
var shozokuID = getObj('team');
var teamNm = "";
var workStr = "";
var yearMonth;
var year;
var month;
var ecb;
var kinmuDate;
try {
    yearMonth = sYearMonth.split("/");
    year = yearMonth[0];
    month = yearMonth[1];
    // 帳票出力用SQL
    strSql.append("SELECT KB.KINMU_DATE , KB.USER_ID , MESAI.KAISHA_ID , CASE WHEN MESAI.GEMBA_ID IS NULL THEN MG2.GEMBA_ID ELSE MESAI.GEMBA_ID END AS GEMBA_ID, MESAI.WORK_TIME , MESAI.TASK_ID");
    strSql.append(", MESAI.TASK_DETAIL , GSAT.WORK_DECIMAL , KB.SYOUNINSHA_ID, KB.START_TIME, KB.END_TIME,KB.KINMU_DAY_WEEK ");
    strSql.append(", MS.SAGYOU_NM,KB.SAGYOU_SUM_TIME_STR,MESAI.START_TIME,MESAI.END_TIME, MT.KAISHA_NM");
    strSql.append(", KB.BIKOU,KB.SHACHOU_SYOUNIN, MU.USER_NAME");
    strSql.append(", CASE WHEN MG.GEMBA_NM IS NULL THEN MG2.GEMBA_NM ELSE MG.GEMBA_NM END AS GEMBA_NM,WEEK(KB.KINMU_DATE, 3) WORK, KB.KINMU_ID, SMU.USER_NAME  FROM KINTAITBL1 KB ");
    strSql.append(" LEFT JOIN T_GEMBA_SAGYOU MESAI ON KB.USER_ID = MESAI.USER_ID AND KB.KINMU_DATE = MESAI.KINMU_DATE AND MESAI.GEMBA_ID = '" + shozokuID + "' ");
    strSql.append(" LEFT JOIN M_SAGYOU_NAIYOU MS ON MS.SAGYOU_ID = MESAI.TASK_ID AND MS.KANRI_ID = MESAI.GEMBA_ID ");
    strSql.append(" LEFT JOIN M_TORIHIKI_KAISHA MT ON MESAI.KAISHA_ID = MT.KAISHA_ID ");
    strSql.append(" LEFT JOIN M_USER MU ON MU.USER_ID = KB.USER_ID ");
    strSql.append(" LEFT JOIN M_USER SMU ON SMU.USER_ID = KB.SHACHOU_ID ");
    strSql.append(" LEFT JOIN M_GEMBA MG ON MESAI.GEMBA_ID = MG.GEMBA_ID ");
    strSql.append(" INNER JOIN M_GEMBA_USER GU ON GU.GEMBA_ID = '" + shozokuID + "' AND GU.USER_ID = MU.USER_ID ");
    strSql.append(" LEFT JOIN M_GEMBA MG2 ON GU.GEMBA_ID = MG2.GEMBA_ID ");
    strSql.append(" LEFT JOIN (SELECT KINMU_DATE, USER_ID, GEMBA_ID, SUM(GSA.WORK_DECIMAL) WORK_DECIMAL ");
    strSql.append(" FROM T_GEMBA_SAGYOU GSA WHERE GSA.GEMBA_ID = '"+shozokuID+"' GROUP BY KINMU_DATE, USER_ID, GEMBA_ID) GSAT ");
    strSql.append(" ON GSAT.KINMU_DATE= KB.KINMU_DATE  AND GSAT.USER_ID = KB.USER_ID ");
    
    //strSql.append(" LEFT JOIN M_SHOZOKU SZ ON SZ.SHOZOKU_ID = MU.SHOZOKU_ID")
    strSql.append(" WHERE 1=1 AND KB.KINMU_DATE BETWEEN GU.START_DATE AND GU.END_DATE ");

    // 出力条件:開始年月
    if (sYearMonth != null && sYearMonth != "") {
        strSql.append(" AND KB.KINMU_YM='" + sYearMonth + "'");
    }
    strSql.append(" ORDER BY KB.USER_ID,KB.KINMU_DATE");
    trace(strSql.toString());
    // 開発ﾌｪｰｽﾞ作業データ取得
    al = exequeryarrylist(strSql.toString());
    excelDataBean = new ExcelDataBean();
    // 開発ﾌｪｰｽﾞ作業データがある場合
    if (al != null && al.size() > 0) {
        // list初期化
        rowDataList = new ArrayList();
        sheeBean = new ExcelDataBean();
        var conditionMap = new HashMap ();
        var index = 1;
        var workDecimal = 0;
        // 社員ID
        userId = al.get(0).get(1);
        workStr = al.get(0).get(21);
        //kinmuDate = al.get(0).get(0);
        var kinmuDate1 = "";
        for (var i = 0; i < al.size(); i++) {
            row1 = new ArrayList();

            row = al.get(i);
            
            // 勤務日1
            row1.add(row.get(0));
            // 勤務日（曜日）2
            row1.add(weekdayName(row.get(11)));
            // 作業区分無し場合3
            if(nullOrBlank(row.get(12))){
                row1.add("");
            } else {
                // 作業区分有場合
                row1.add(row.get(12));
            }
            // 作業内容無し場合4
            if(nullOrBlank(row.get(6))){
                row1.add("");
            } else {
                // 作業内容有場合
                row1.add(row.get(6));
            }
            // 作業時間無し場合5
            if(nullOrBlank(row.get(4))){
                row1.add("");
            } else {
                // 作業時間有場合
                row1.add(row.get(4));
            }
            // 	作業場所無し場合6
            if(nullOrBlank(row.get(16))){
                row1.add("");
            } else {
                // 作業場所有場合
                row1.add(row.get(16));
            }
            // 勤務開始時間無し場合7
            if(nullOrBlank(row.get(9))){
                row1.add("");
            } else {
                // 勤務開始時間有場合
                row1.add(row.get(9));
            }
            // 勤務終了時間無し場合8
            if(nullOrBlank(row.get(10))){
                row1.add("");
            } else {
                // 勤務終了時間有場合
                row1.add(row.get(10));
            }
            // 作業詳細時間合計無し場合9
            if(nullOrBlank(row.get(7))){
                row1.add("");
            } else {
                // 作業詳細時間合計有場合
                row1.add(getHHmm(gethourminute(row.get(7))));
            }
            // BIKOU コメント無し場合10
            if(nullOrBlank(row.get(17))){
                row1.add("");
            } else {
                // BIKOU コメント有場合
                row1.add(row.get(17));
            }
            // 承認者無し場合11
            if(nullOrBlank(row.get(23))){
                row1.add("");
            } else {
                // 承認者有場合
                row1.add(row.get(23));
            }
            // 承認者無し場合12
            if(nullOrBlank(row.get(23))){
                row1.add("");
            } else {
                // 承認者有場合
                row1.add(row.get(23));
            }
            if (String(userId) != String(row.get(1))) {
                
                sheeBean.addItem("K3",year);
                sheeBean.addItem("K4",month);
                sheeBean.addItem("K5",teamNm);
                sheeBean.addItem("K6",userNm);
                
                //帳票ヘッダを複製
                sheeBean.addCopyItem("A1", "A8");
                //固定セルに値を設定
                sheeBean.addItem("A2", "XXXXXﾌﾟﾛｼﾞｪｸﾄ 作業報告書");
                // 勤務日 セルを日付タイプに設定
                conditionMap.put(0, new ExcelCellBean(0, 0, "date"));
                // 勤務日（曜日） セルを文字列タイプに設定
                conditionMap.put(1, new ExcelCellBean(0, 1, "string"));
                // 作業区分 セルを文字列タイプに設定
                conditionMap.put(2, new ExcelCellBean(0, 2, "string"));
                // 作業内容 セルを文字列タイプに設定
                conditionMap.put(3, new ExcelCellBean(0, 3, "string"));
                // 作業時間 セルを時刻タイプに設定
                conditionMap.put(4, new ExcelCellBean(0, 4, "time"));
                // 作業場所 セルを文字列タイプに設定
                ecb = new ExcelCellBean(0, 5, "string");
                conditionMap.put(5, ecb);
                // 始業時刻 セルを時刻タイプに設定
                conditionMap.put(6, new ExcelCellBean(0, 6, "time"));
                // 終業時刻 セルを時刻タイプに設定
                conditionMap.put(7, new ExcelCellBean(0, 7, "time"));
                // 勤務時間 セルを時刻タイプに設定
                conditionMap.put(8, new ExcelCellBean(0, 8, "time"));
                // 報告事項 セルを文字列タイプに設定
                conditionMap.put(9, new ExcelCellBean(0, 9, "string"));
                // 承認者 セルを文字列タイプに設定
                conditionMap.put(10, new ExcelCellBean(0, 10, "string"));
                // 隠す列（週小計） セルを文字列タイプに設定
                conditionMap.put(11, new ExcelCellBean(0, 11, "string"));

                // データとデータタイプ定義をExcelテーブルの指定開始セルに指定、同時にExcelテーブル定義のListBeanを取得
                var listBean = sheeBean.addList("A9", rowDataList, conditionMap);
                
                // 小計計算を定義
                var listcount = listBean.setCount("A10", "A10", "M9");
                
                // 月日セルが同じ場合、上下セルをマージする
                listBean.setBreak(0, 0);
                // 月日セルが同じ場合、曜日の上下セルをマージする
                listBean.setBreak(1, 0);
                // 月日セルが同じ場合、始業時刻の上下セルをマージする
                listBean.setBreak(6, 0);
                // 月日セルが同じ場合、終業時刻の上下セルをマージする
                listBean.setBreak(7, 0);
                // 月日セルが同じ場合、勤務時間の上下セルをマージする
                listBean.setBreak(8, 0);
                // 月日セルが同じ場合、報告事項の上下セルをマージする
                listBean.setBreak(9, 0);
                // 月日セルが同じ場合、承認者の上下セルをマージする
                listBean.setBreak(10, 0);
                // 月日セルが同じ場合、週小計の上下セルをマージする
                listBean.setBreak(11, 0);                
                // 月次合計行の合計を定義
                sheeBean.addRelativeCopyItem("A11", "A11", 0, 0, false, true, true);
                // UserIDとUser名をシート名に設定
                excelDataBean.addSheet(userId + " " + userNm, sheeBean);
                
                sheeBean = new ExcelDataBean();
                conditionMap = new HashMap();
                rowDataList = new ArrayList();
                index = 0;
            }
            // 隠す列（週小計）
            if(row.get(21) != workStr){
                index++;
                // 隠す列（週小計）
                row1.add(String(index));
            } else {
                // 隠す列（週小計）
                row1.add(String(index));
            }
            rowDataList.add(row1);
            userId = row.get(1);
            userNm = row.get(19);
            workStr = row.get(21);
            teamNm = row.get(20);
            
        }
        
        //帳票ヘッダを複製
        sheeBean.addCopyItem("A1", "A8");
        
        sheeBean.addItem("K3",year);
        sheeBean.addItem("K4",month);
        sheeBean.addItem("K5",teamNm);
        sheeBean.addItem("K6",userNm);
        
        //固定セルに値を設定
        sheeBean.addItem("A2", "XXXXXﾌﾟﾛｼﾞｪｸﾄ 作業報告書");
        // 勤務日 セルを日付タイプに設定
        conditionMap.put(0, new ExcelCellBean(0, 0, "date"));
        // 勤務日（曜日） セルを文字列タイプに設定
        conditionMap.put(1, new ExcelCellBean(0, 1, "string"));
        // 作業区分 セルを文字列タイプに設定
        conditionMap.put(2, new ExcelCellBean(0, 2, "string"));
        // 作業内容 セルを文字列タイプに設定
        conditionMap.put(3, new ExcelCellBean(0, 3, "string"));
        // 作業時間 セルを時刻タイプに設定
        conditionMap.put(4, new ExcelCellBean(0, 4, "time"));
        // 作業場所 セルを文字列タイプに設定
        ecb = new ExcelCellBean(0, 5, "string");
        conditionMap.put(5, ecb);
        // 始業時刻 セルを時刻タイプに設定
        conditionMap.put(6, new ExcelCellBean(0, 6, "time"));
        // 終業時刻 セルを時刻タイプに設定
        conditionMap.put(7, new ExcelCellBean(0, 7, "time"));
        // 勤務時間 セルを時刻タイプに設定
        conditionMap.put(8, new ExcelCellBean(0, 8, "time"));
        // 報告事項 セルを文字列タイプに設定
        conditionMap.put(9, new ExcelCellBean(0, 9, "String"));
        // 承認者 セルを文字列タイプに設定
        conditionMap.put(10, new ExcelCellBean(0, 10, "string"));
        // 隠す列（週小計） セルを文字列タイプに設定
        conditionMap.put(11, new ExcelCellBean(0, 11, "string"));

        // データとデータタイプ定義をExcelテーブルの指定開始セルに指定、同時にExcelテーブル定義のListBeanを取得
        var listBean = sheeBean.addList("A9", rowDataList, conditionMap);
        // 小計計算を定義
        var listcount = listBean.setCount("A10", "A10", "M9");
        // 月日セルが同じ場合、上下セルをマージする
        listBean.setBreak(0, 0);
        // 月日セルが同じ場合、曜日の上下セルをマージする
        listBean.setBreak(1, 0);
        // 月日セルが同じ場合、始業時刻の上下セルをマージする
        listBean.setBreak(6, 0);
        // 月日セルが同じ場合、終業時刻の上下セルをマージする
        listBean.setBreak(7, 0);
        // 月日セルが同じ場合、勤務時間の上下セルをマージする
        listBean.setBreak(8, 0);
        // 月日セルが同じ場合、報告事項の上下セルをマージする
        listBean.setBreak(9, 0);
        // 月日セルが同じ場合、承認者の上下セルをマージする
        listBean.setBreak(10, 0);
        // 月日セルが同じ場合、週小計の上下セルをマージする
        listBean.setBreak(11, 0);
        
        // 月次合計行の合計を定義
        sheeBean.addRelativeCopyItem("A11", "A11", 0, 0, false, true, true);
        // UserIDとUser名をシート名に設定
        excelDataBean.addSheet(userId + " " + userNm, sheeBean);
        // 印刷幅は1ページ
        excelDataBean.setFitWidth(1);
        // 印刷高さは1ページ
        excelDataBean.setFitHeight(1);
        //縦印刷
        excelDataBean.setLandscape(false);
        // footerを設定
        excelDataBean.setFooterCenter("%page% / %numPages%");
        // 水平中央に設定
        excelDataBean.setHorizontallyCenter(true);
        // 定義内容をexceldatabeanに追加
        excelPrint.setExcelDataBean(excelDataBean);
        
        // 帳票テンプレート番号を設定
        excelPrint.setTemplateID("2348");
        // 帳票を印刷
        JsonObj = excelPrint.print("2348");
        //ActStr = "alert('正常に出力しました。');";
    } else {
        // データが存在しない場合、メッセージを表示
        excelDataBean.setErrorMsg("該当データは存在しません。");
        excelPrint.setTemplateID("2348");
        excelPrint.setExcelDataBean(excelDataBean);
        JsonObj = excelPrint.print("2348");
    }

} catch (e) {
    trace(e);
    ActStr = "alert('出力時エラーが発生しました。システム管理者にご連絡ください。[" + e  + "]');";
}

function weekdayName(weekday){
    
    if(weekday == "0"){
        return "日";
    } else if(weekday == "1") {
        return "月";
    } else if(weekday == "2") {
        return "火";
    } else if(weekday == "3") {
        return "水";
    } else if(weekday == "4") {
        return "木";
    } else if(weekday == "5") {
        return "金";
    } else if(weekday == "6") {
        return "土";
    }
}

function gethourminute(minute) {

    if (minute != null && minute != "") {
        minute = parseInt(minute*60,10);
    } else {
        minute = 0;
    }
    return minute;
}
function getHHmm(minutes) {
    if (minutes == "") {
      return "";
    }
    var hh = Math.floor(minutes/60);

    var mm = minutes - hh * 60;
    var mmlbl = "";
    if (mm < 10) {
      mmlbl = "0" + mm;
    } else {
      mmlbl = mm;
    }
    if(hh < 10){
       hh = "0" + hh; 
    }
    return hh + ":" + mmlbl;
}
