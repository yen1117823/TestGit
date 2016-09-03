<%
'//
'// Copyright (c) 2012 BrainLab Inc., All rights reserved.
'// $Id$
'//
'----------------------------------------------------------------
'履歴：
'		2009.03.04　n.hiraoka　カート機能
'					ORDER_IDはRequestから取らずにカートから取る
'		2009.06.16	ukaji		プログラム毎のDBオープン、クローズ対応
'// 履歴		：2009.08.19 yamaguchik no-chace をsystem.envへ移行
'// 履歴		：2016.08.24 YenHTP 対応
'----------------------------------------------------------------
%>
<!--#include file="../include/sessionCn_Open.inc"-->
<!--#include file="../include/SessionDetect.inc"-->
<!--#include file="../include/cp2ctrl.inc"-->
<!--#include file="../include/XlsCreate_OrderSheet_xlsx.inc"-->

<%
response.buffer = true

err_msg = ""

Set mRec = Server.CreateObject("ADODB.Recordset")

'//画面反映用フィールドデータ保持用ディクショナリ
Set V_ORD515 = Server.CreateObject("Scripting.Dictionary")
'//ディクショナリ連想配列の比較モードをバイナリにする
V_ORD515.Comparemode=1

'=============================
'--GET

'---- 2009.03.04 n.hiraoka ----
Chk_ORDER_ID = Request("chk")
Career_ID     = Request("CAREER_ID_T")
T_ORDER_ID    = Request("ORDER_ID_T")

if chkEmp(Chk_ORDER_ID) and chkEmp(T_ORDER_ID) then
	Set mRecCart = Server.CreateObject("ADODB.Recordset")

	wSQL = "select order_id from trnOrder_cart where user_id =" & session("userid")
	mRecCart.open wSQL, Cn,0,1

		if mRecCart.eof then
			err_msg = "カートが空です。\n " & V_err
		end if
		Do until mRecCart.eof
			Chk_ORDER_ID = Chk_ORDER_ID & "," & mRecCart(0)
			mRecCart.movenext
		Loop

	mRecCart.close
	set  mRecCart = nothing

	Chk_ORDER_ID = Mid(Chk_ORDER_ID,2)
End if

'---- 2009.03.04 n.hiraoka ----
'---- 2016.08.24 YenHTP.Begin ----
Dim wSQLCheckSTT
	wSQLCheckSTT =""
Dim mRecCheckSTT

If Chk_ORDER_ID <> "" Then 
    Set mRecCheckSTT = Server.CreateObject("ADODB.Recordset")
    wSQLCheckSTT = "Select TOR.ORDER_ID,TOR.POSITIONNAME,REGSTATUS_ID from TRNORDER TOR "
    wSQLCheckSTT = wSQLCheckSTT & " Where TOR.ORDER_ID IN (" & Chk_ORDER_ID & ")"
    mRecCheckSTT.open wSQLCheckSTT, Cn,0,1	
    Chk_ORDER_ID = ""   
    Do until mRecCheckSTT.eof
        '//求人情報
        If mRecCheckSTT("REGSTATUS_ID") = "3" Then
	        '// 募集終了
	        err_msg = err_msg & mRecCheckSTT("ORDER_ID") & ":" & mRecCheckSTT("POSITIONNAME") & "\n"
        else
	        Chk_ORDER_ID = Chk_ORDER_ID & "," & mRecCheckSTT("ORDER_ID")
        End If
        mRecCheckSTT.movenext
    Loop
    mRecCheckSTT.close
    If Len(err_msg) > 0 Then
        err_msg = "下記求人が募集終了のため送信できません。\n 再度、選択して下さい。\n" & err_msg            
    End If
    If Chk_ORDER_ID <> "" Then
        Chk_ORDER_ID = mid(Chk_ORDER_ID,2)
    Else
        err_msg = err_msg & "\n求人票送信が可能な求人が存在しません。\n 再度、選択して下さい。"
    End if
    If err_msg <> "" Then
        '// 画面にエラーを表示し、閉じる
%>
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
<script type="text/javascript">
    alert('<%=err_msg%>');
    self.close();
</script>
<%
        Response.End
    end if
End If
'---- 2016.08.24 YenHTP.End ----
h_AddFlg      = Request("h_AddFlg")
h_CAREER_ID_L = Request("h_CAREER_ID_L")
Career_ID     = Request("CAREER_ID_T")
T_ORDER_ID    = Request("ORDER_ID_T")


'response.write Career_ID & "<br>"
'response.write T_ORDER_ID & "<br>"
'response.write h_AddFlg & "<br>"
V_ORD515("CAREER_ID_L") = h_CAREER_ID_L
V_ORD515("CAREER_ID") = CAREER_ID
If (Chk_ORDER_ID = "") Then
	Chk_ORDER_ID = T_ORDER_ID
End If

'=============================
'---- 2016.08.24 YenHTP.Begin ----
h_unCheckOrder = Request("h_unCheckOrder")
h_unCheckOrder = Mid(h_unCheckOrder,2)
Dim ck_ORDER_ID_515
    ck_ORDER_ID_515 = Request("ck_ORDER_ID_515")
    ck_ORDER_ID_515 = Mid(ck_ORDER_ID_515,2)

Set V_ck_ORDER_ID_515 = Server.CreateObject("Scripting.Dictionary")
If (h_AddFlg = "Y") and (Career_ID <> "") and (T_ORDER_ID <> "") Then
     Dim MYPAGEGROUP_NAME_T
     Dim MYPAGEGROUP_COMMENT_T
     Dim DISPORDER
     Dim RCM_FLG
         
     Set List_DISPORDER = Server.CreateObject("Scripting.Dictionary")
     List_DISPORDER.Comparemode=1
         
     Set List_RCM_FLG = Server.CreateObject("Scripting.Dictionary")
     List_RCM_FLG.Comparemode=1

     MYPAGEGROUP_NAME_T =""
     MYPAGEGROUP_COMMENT_T = ""
     MYPAGEGROUP_NAME_T = Request("MYPAGEGROUP_NAME_T")
     MYPAGEGROUP_COMMENT_T = Request("MYPAGEGROUP_COMMENT_T")
     
     DISPORDER = Request("h_DISPORDER")  
     DISPORDER = Mid(DISPORDER,2)   
     List_DISPORDER = Split(DISPORDER,",",-1,1)

     RCM_FLG = Request("h_RCM_FLG")
     RCM_FLG = Mid(RCM_FLG,2)  
     List_RCM_FLG = Split(RCM_FLG,",",-1,1)
'---- 2016.08.24 YenHTP.End ----
	'追加ボタン押下後の場合
	'// Response.Write "Career_ID =" & Career_ID & "<BR>"
	'//V_Order_ID = Split(T_ORDER_ID,",",-1,1)
'---- 2016.08.24 YenHTP.Begin ----	
	V_ck_ORDER_ID_515 = Split(ck_ORDER_ID_515,",",-1,1)
	If Len(Career_ID) > 0 Then
		For l= 0 To Ubound(V_ck_ORDER_ID_515)
			'// Response.Write l & "=" & V_Order_ID(l) & "<HR>"
			Memo="一括進捗追加機能より"
            
			'//既存進捗を削除
			'wSQL = "delete  from trnShintyoku  "
			'wSQL = wSQL  & " where Career_ID = " & Career_ID & " and Order_ID = " & V_Order_ID(l)
            
            'YenHTP SQL文21
            wSQL = "delete  from TRNSHINTYOKU  "
			wSQL = wSQL  & " where Career_ID = " & Career_ID & " and Order_ID = " & V_ck_ORDER_ID_515(l) & " and PROGRESSFINISH_ID = 1"
'---- 2016.08.24 YenHTP.End ----

'// 2009/06/16 ukaji プログラム毎のDBオープン、クローズ対応
'			Session("db").execute wSQL
			Cn.execute wSQL

			'//新規SEQ－取得
			wSQL = "select SHINTYOKU_SEQ.nextval from dual"
'// 2009/06/16 ukaji プログラム毎のDBオープン、クローズ対応
'			mRec.open wSQL, Session("db"),0,1
'---- 2016.08.24 YenHTP.Begin ----
            Set mRec = Server.CreateObject("ADODB.Recordset")
			mRec.open wSQL, Cn,0,1
			PROGRESS_ID = mRec(0)
			mRec.close()
            Set mRec = Nothing

            'YenHTP SQL文17
            Dim MYPAGEGROUP_ID
            MYPAGEGROUP_ID = ""
            wSQL = "select MYPAGEGROUP_SEQ.nextval from dual"
            Set mRec = Server.CreateObject("ADODB.Recordset")
            mRec.open wSQL, Cn,0,1
            MYPAGEGROUP_ID = mRec(0)
            mRec.close()
            Set mRec = Nothing
			'//データ作成
			'//TRNSHINTYOKU
			'wSQL = "Insert into TRNSHINTYOKU(" & _
			'		"PROGRESS_ID,CAREER_ID,Order_ID,STATUS_ID,MEMO,LASTPROGRESS_DATE) " & _
			'		"values(" & _
			'		PROGRESS_ID & "," & Career_ID & "," & V_Order_ID(l) & ",11,'" & Memo & "',sysdate)"

            'YenHTP SQL文31
            wSQL = "Insert into TRNSHINTYOKU(" & _
					"PROGRESS_ID,CAREER_ID,Order_ID,STATUS_ID,MEMO,LASTPROGRESS_DATE,LASTRESULT_ID) " & _
					"values(" & PROGRESS_ID & "," & Career_ID & "," & V_ck_ORDER_ID_515(l) & ",11,'NULL',sysdate,32768)"
'---- 2016.08.24 YenHTP ----

'// 2009/06/16 ukaji プログラム毎のDBオープン、クローズ対応
'			Session("db").execute wSQL
			Cn.Execute wSQL

			'//TRNSHINTYOKUSUB
			'wSQL = "Insert into trnShintyokuSub(" & _
			'		"PROGRESS_ID,PROGRESS_ID_Sub,PROGRESS_DATE,Result_ID,UPDATEUser_ID,Memo, " & _
			'		"CAREER_CHARGETEAM_ID,CAREER_CHARGE_ID,ORDER_CHARGETEAM_ID,ORDER_CHARGE_ID) " & _
			'		"values(" & PROGRESS_ID & ",1,sysdate,32768," & Session("UserID") & ",'" & Memo & "'," & _
			'		"nvl((select CHARGETEAM_ID from TRNCAREER where CAREER_ID = " & Career_ID & "), 0)," & _
			'		"nvl((select CHARGE_ID from TRNCAREER where CAREER_ID = " & Career_ID & "), 0)," & _
			'		"nvl((select t1.CHARGETEAM_ID from TRNCLIENTSUB t1 where exists(" & _
			'		"select * from TRNORDER t2 where t2.ORDER_ID = " & V_Order_ID(l) & _
			'		" and t1.CLIENT_ID = t2.CLIENT_ID and t1.CLIENTSUB_ID = t2.CLIENTSUB_ID)), 0)," & _
			'		"nvl((select t1.CHARGE_ID from TRNCLIENTSUB t1 where exists(" & _
			'		"select * from TRNORDER t2 where t2.ORDER_ID = " & V_Order_ID(l) & _
			'		" and t1.CLIENT_ID = t2.CLIENT_ID and t1.CLIENTSUB_ID = t2.CLIENTSUB_ID)), 0))"

'---- 2016.08.24 YenHTP.Begin ----
            '//TRNSHINTYOKUSUB SQL文32
			wSQL = "Insert into trnShintyokuSub(" & _
					"PROGRESS_ID,PROGRESS_ID_Sub,PROGRESS_DATE,Result_ID,UPDATEUser_ID, " & _
					"CAREER_CHARGETEAM_ID,CAREER_CHARGE_ID,ORDER_CHARGETEAM_ID,ORDER_CHARGE_ID) " & _
					"values(" & PROGRESS_ID & ",1,sysdate,32768," & Session("UserID") & "," & _
					"nvl((select CHARGETEAM_ID from TRNCAREER where CAREER_ID = " & Career_ID & "), 0)," & _
					"nvl((select CHARGE_ID from TRNCAREER where CAREER_ID = " & Career_ID & "), 0)," & _
					"nvl((select t1.CHARGETEAM_ID from TRNCLIENTSUB t1 where exists(" & _
					"select * from TRNORDER t2 where t2.ORDER_ID = " & V_ck_ORDER_ID_515(l) & _
					" and t1.CLIENT_ID = t2.CLIENT_ID and t1.CLIENTSUB_ID = t2.CLIENTSUB_ID)), 0)," & _
					"nvl((select t1.CHARGE_ID from TRNCLIENTSUB t1 where exists(" & _
					"select * from TRNORDER t2 where t2.ORDER_ID = " & V_ck_ORDER_ID_515(l) & _
					" and t1.CLIENT_ID = t2.CLIENT_ID and t1.CLIENTSUB_ID = t2.CLIENTSUB_ID)), 0))"
'---- 2016.08.24 YenHTP.End ----

'// 2009/06/16 ukaji プログラム毎のDBオープン、クローズ対応
'			Session("db").execute wSQL
			Cn.Execute wSQL

'---- 2016.08.24 YenHTP.Begin ----
            '//TRNMYPAGEGROUP SQL文33
            wSQL = "Insert into TRNMYPAGEGROUP (" & _
                    "MYPAGEGROUP_ID,MYPAGEGROUP_NAME,MYPAGEGROUP_COMMENT,PROGRESS_ID,CAREER_ID," & _
                    "ORDER_ID,DISPORDER,RCM_FLG,INSERT_DATE,UPDATE_DATE,UPDATEUSR_ID,DELETE_DATE)" & _
                    "values(" & MYPAGEGROUP_ID & ",'" & MYPAGEGROUP_NAME_T &"','" & MYPAGEGROUP_COMMENT_T & "'," & _
                    PROGRESS_ID &"," & Career_ID & "," & V_ck_ORDER_ID_515(l) & "," & List_DISPORDER(l) & "," & List_RCM_FLG(l) & "," & _
                    "sysdate, sysdate, " & Session("UserID") & ", NULL)"
            Cn.Execute wSQL
			'PROGRESSCHECK_IDを入れる処理を追加 080206 eno
			'wSQL = "UPDATE TRNORDER SET PROGRESSCHECK_ID = 1,LASTPROGRESS_DATE = '" & now() & "' Where ORDER_ID = " & V_Order_ID(l)
			wSQL = "UPDATE TRNORDER SET PROGRESSCHECK_ID = 1,LASTPROGRESS_DATE = '" & now() & "' Where ORDER_ID = " & V_ck_ORDER_ID_515(l)
'---- 2016.08.24 YenHTP.End ----
'// 2009/06/16 ukaji プログラム毎のDBオープン、クローズ対応
'			Session("db").Execute wSQL
			Cn.Execute wSQL
			wSQL = "UPDATE TRNCAREER SET PROGRESSCHECK_ID = 1 Where CAREER_ID = " & CAREER_ID
'// 2009/06/16 ukaji プログラム毎のDBオープン、クローズ対応
'			Session("db").Execute wSQL
			Cn.Execute wSQL

			err_msg = "一括追加完了しました。"
		Next
	End If
End if

set mRec = nothing

Set mRecOrd = Server.CreateObject("ADODB.Recordset")

	'//Unified_Checker
'	wSQL = "SELECT NAME, IME || NOTNULL || ONLYNUMERIC || ONLYALPHANUMERIC || NOTHALFSIZEKANA || NOTSPACE || ONLYEMAIL || NOTRULECHAR || NOTFULLSIZEKANA || ' ' || lpad(BEAMS_MIN, 13, '0') ||  ' ' || lpad(BEAMS_MAX, 13, '0') ||  ' ' || lpad(BEAMS_SHORTEST, 13, '0') || ' ' || lpad(BEAMS_LONGEST, 13, '0') AS ID FROM MSTUNIFIED_CHECKER "
'	wSQL = wSQL & " Where FILENAME = '5-15'"

'// 2009/06/16 ukaji プログラム毎のDBオープン、クローズ対応
'	mRecOrd.open wSQL, Session("db"),0,1
'	mRecOrd.open wSQL, Cn,0,1
'
'	Do until mRecOrd.eof
'		V_Ord515.Item(mRecOrd("NAME")  & "_ID") = mRecOrd("ID")
'		V_Ord515.Item("ALL_UNI_FIELD") = V_Ord515.Item("ALL_UNI_FIELD") & "," & mRecOrd("NAME")
'		mRecOrd.movenext
'	Loop
'	mRecOrd.close

	'//TAG作成
'	W_FIELDNAME = split(V_Ord515("ALL_UNI_FIELD"),",",-1,1)
'	For X= LBound(W_FIELDNAME) To UBound(W_FIELDNAME)
'		if Len(W_FIELDNAME(x)) > 0 Then
'			V_Ord515.Item(W_FIELDNAME(x) & "_TAG") = Convert_tag(W_FIELDNAME(x),V_Ord515(W_FIELDNAME(x) & "_ID"),"")
'		End If
'	Next

set mRecOrd = nothing
'---- 2016.08.24 YenHTP.Begin ----
Dim IsJobPosting
	IsJobPosting = False
    IsJobPosting = Request("IsJobPosting")
'//Arr status progress
    Set V_ORD_STATUS_515 = Server.CreateObject("Scripting.Dictionary")    
    V_ORD_STATUS_515.Comparemode=1
    V_ORD_STATUS_515(16) = "RAから依頼中"
    V_ORD_STATUS_515(11) = "応募意思確認中（求人）"
    V_ORD_STATUS_515(12) = "書類提出待ち"
    V_ORD_STATUS_515(1) = "書類結果待ち"
    V_ORD_STATUS_515(2) = "面接設定中"
    V_ORD_STATUS_515(3) = "面接結果待ち"
    V_ORD_STATUS_515(4) = "入社意思確認中（内定）"
    V_ORD_STATUS_515(6) = "入社確認待ち"
    V_ORD_STATUS_515(14) = "請求処理中"
    V_ORD_STATUS_515(8) = "企業へNG連絡中"
    V_ORD_STATUS_515(7) = "求職者へNG連絡中"
    V_ORD_STATUS_515(17) = "NG終了"  
    V_ORD_STATUS_515(9) = "辞退終了"
    V_ORD_STATUS_515(10) = "完了"
'---- 2016.08.24 YenHTP.End ----
%>
<!--#include file="./5-15.htm"-->
<!--#include file="../include/sessionCn_Close.inc"-->
