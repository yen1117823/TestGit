<%
'//
'// Copyright (c) 2012 BrainLab Inc., All rights reserved.
'// $Id$
'//
'----------------------------------------------------------------
'�����F
'		2009.03.04�@n.hiraoka�@�J�[�g�@�\
'					ORDER_ID��Request�����炸�ɃJ�[�g������
'		2009.06.16	ukaji		�v���O��������DB�I�[�v���A�N���[�Y�Ή�
'// ����		�F2009.08.19 yamaguchik no-chace ��system.env�ֈڍs
'// ����		�F2016.08.24 YenHTP �Ή�
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

'//��ʔ��f�p�t�B�[���h�f�[�^�ێ��p�f�B�N�V���i��
Set V_ORD515 = Server.CreateObject("Scripting.Dictionary")
'//�f�B�N�V���i���A�z�z��̔�r���[�h���o�C�i���ɂ���
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
			err_msg = "�J�[�g����ł��B\n " & V_err
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
        '//���l���
        If mRecCheckSTT("REGSTATUS_ID") = "3" Then
	        '// ��W�I��
	        err_msg = err_msg & mRecCheckSTT("ORDER_ID") & ":" & mRecCheckSTT("POSITIONNAME") & "\n"
        else
	        Chk_ORDER_ID = Chk_ORDER_ID & "," & mRecCheckSTT("ORDER_ID")
        End If
        mRecCheckSTT.movenext
    Loop
    mRecCheckSTT.close
    If Len(err_msg) > 0 Then
        err_msg = "���L���l����W�I���̂��ߑ��M�ł��܂���B\n �ēx�A�I�����ĉ������B\n" & err_msg            
    End If
    If Chk_ORDER_ID <> "" Then
        Chk_ORDER_ID = mid(Chk_ORDER_ID,2)
    Else
        err_msg = err_msg & "\n���l�[���M���\�ȋ��l�����݂��܂���B\n �ēx�A�I�����ĉ������B"
    End if
    If err_msg <> "" Then
        '// ��ʂɃG���[��\�����A����
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
	'�ǉ��{�^��������̏ꍇ
	'// Response.Write "Career_ID =" & Career_ID & "<BR>"
	'//V_Order_ID = Split(T_ORDER_ID,",",-1,1)
'---- 2016.08.24 YenHTP.Begin ----	
	V_ck_ORDER_ID_515 = Split(ck_ORDER_ID_515,",",-1,1)
	If Len(Career_ID) > 0 Then
		For l= 0 To Ubound(V_ck_ORDER_ID_515)
			'// Response.Write l & "=" & V_Order_ID(l) & "<HR>"
			Memo="�ꊇ�i���ǉ��@�\���"
            
			'//�����i�����폜
			'wSQL = "delete  from trnShintyoku  "
			'wSQL = wSQL  & " where Career_ID = " & Career_ID & " and Order_ID = " & V_Order_ID(l)
            
            'YenHTP SQL��21
            wSQL = "delete  from TRNSHINTYOKU  "
			wSQL = wSQL  & " where Career_ID = " & Career_ID & " and Order_ID = " & V_ck_ORDER_ID_515(l) & " and PROGRESSFINISH_ID = 1"
'---- 2016.08.24 YenHTP.End ----

'// 2009/06/16 ukaji �v���O��������DB�I�[�v���A�N���[�Y�Ή�
'			Session("db").execute wSQL
			Cn.execute wSQL

			'//�V�KSEQ�|�擾
			wSQL = "select SHINTYOKU_SEQ.nextval from dual"
'// 2009/06/16 ukaji �v���O��������DB�I�[�v���A�N���[�Y�Ή�
'			mRec.open wSQL, Session("db"),0,1
'---- 2016.08.24 YenHTP.Begin ----
            Set mRec = Server.CreateObject("ADODB.Recordset")
			mRec.open wSQL, Cn,0,1
			PROGRESS_ID = mRec(0)
			mRec.close()
            Set mRec = Nothing

            'YenHTP SQL��17
            Dim MYPAGEGROUP_ID
            MYPAGEGROUP_ID = ""
            wSQL = "select MYPAGEGROUP_SEQ.nextval from dual"
            Set mRec = Server.CreateObject("ADODB.Recordset")
            mRec.open wSQL, Cn,0,1
            MYPAGEGROUP_ID = mRec(0)
            mRec.close()
            Set mRec = Nothing
			'//�f�[�^�쐬
			'//TRNSHINTYOKU
			'wSQL = "Insert into TRNSHINTYOKU(" & _
			'		"PROGRESS_ID,CAREER_ID,Order_ID,STATUS_ID,MEMO,LASTPROGRESS_DATE) " & _
			'		"values(" & _
			'		PROGRESS_ID & "," & Career_ID & "," & V_Order_ID(l) & ",11,'" & Memo & "',sysdate)"

            'YenHTP SQL��31
            wSQL = "Insert into TRNSHINTYOKU(" & _
					"PROGRESS_ID,CAREER_ID,Order_ID,STATUS_ID,MEMO,LASTPROGRESS_DATE,LASTRESULT_ID) " & _
					"values(" & PROGRESS_ID & "," & Career_ID & "," & V_ck_ORDER_ID_515(l) & ",11,'NULL',sysdate,32768)"
'---- 2016.08.24 YenHTP ----

'// 2009/06/16 ukaji �v���O��������DB�I�[�v���A�N���[�Y�Ή�
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
            '//TRNSHINTYOKUSUB SQL��32
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

'// 2009/06/16 ukaji �v���O��������DB�I�[�v���A�N���[�Y�Ή�
'			Session("db").execute wSQL
			Cn.Execute wSQL

'---- 2016.08.24 YenHTP.Begin ----
            '//TRNMYPAGEGROUP SQL��33
            wSQL = "Insert into TRNMYPAGEGROUP (" & _
                    "MYPAGEGROUP_ID,MYPAGEGROUP_NAME,MYPAGEGROUP_COMMENT,PROGRESS_ID,CAREER_ID," & _
                    "ORDER_ID,DISPORDER,RCM_FLG,INSERT_DATE,UPDATE_DATE,UPDATEUSR_ID,DELETE_DATE)" & _
                    "values(" & MYPAGEGROUP_ID & ",'" & MYPAGEGROUP_NAME_T &"','" & MYPAGEGROUP_COMMENT_T & "'," & _
                    PROGRESS_ID &"," & Career_ID & "," & V_ck_ORDER_ID_515(l) & "," & List_DISPORDER(l) & "," & List_RCM_FLG(l) & "," & _
                    "sysdate, sysdate, " & Session("UserID") & ", NULL)"
            Cn.Execute wSQL
			'PROGRESSCHECK_ID�����鏈����ǉ� 080206 eno
			'wSQL = "UPDATE TRNORDER SET PROGRESSCHECK_ID = 1,LASTPROGRESS_DATE = '" & now() & "' Where ORDER_ID = " & V_Order_ID(l)
			wSQL = "UPDATE TRNORDER SET PROGRESSCHECK_ID = 1,LASTPROGRESS_DATE = '" & now() & "' Where ORDER_ID = " & V_ck_ORDER_ID_515(l)
'---- 2016.08.24 YenHTP.End ----
'// 2009/06/16 ukaji �v���O��������DB�I�[�v���A�N���[�Y�Ή�
'			Session("db").Execute wSQL
			Cn.Execute wSQL
			wSQL = "UPDATE TRNCAREER SET PROGRESSCHECK_ID = 1 Where CAREER_ID = " & CAREER_ID
'// 2009/06/16 ukaji �v���O��������DB�I�[�v���A�N���[�Y�Ή�
'			Session("db").Execute wSQL
			Cn.Execute wSQL

			err_msg = "�ꊇ�ǉ��������܂����B"
		Next
	End If
End if

set mRec = nothing

Set mRecOrd = Server.CreateObject("ADODB.Recordset")

	'//Unified_Checker
'	wSQL = "SELECT NAME, IME || NOTNULL || ONLYNUMERIC || ONLYALPHANUMERIC || NOTHALFSIZEKANA || NOTSPACE || ONLYEMAIL || NOTRULECHAR || NOTFULLSIZEKANA || ' ' || lpad(BEAMS_MIN, 13, '0') ||  ' ' || lpad(BEAMS_MAX, 13, '0') ||  ' ' || lpad(BEAMS_SHORTEST, 13, '0') || ' ' || lpad(BEAMS_LONGEST, 13, '0') AS ID FROM MSTUNIFIED_CHECKER "
'	wSQL = wSQL & " Where FILENAME = '5-15'"

'// 2009/06/16 ukaji �v���O��������DB�I�[�v���A�N���[�Y�Ή�
'	mRecOrd.open wSQL, Session("db"),0,1
'	mRecOrd.open wSQL, Cn,0,1
'
'	Do until mRecOrd.eof
'		V_Ord515.Item(mRecOrd("NAME")  & "_ID") = mRecOrd("ID")
'		V_Ord515.Item("ALL_UNI_FIELD") = V_Ord515.Item("ALL_UNI_FIELD") & "," & mRecOrd("NAME")
'		mRecOrd.movenext
'	Loop
'	mRecOrd.close

	'//TAG�쐬
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
    V_ORD_STATUS_515(16) = "RA����˗���"
    V_ORD_STATUS_515(11) = "����ӎv�m�F���i���l�j"
    V_ORD_STATUS_515(12) = "���ޒ�o�҂�"
    V_ORD_STATUS_515(1) = "���ތ��ʑ҂�"
    V_ORD_STATUS_515(2) = "�ʐڐݒ蒆"
    V_ORD_STATUS_515(3) = "�ʐڌ��ʑ҂�"
    V_ORD_STATUS_515(4) = "���Јӎv�m�F���i����j"
    V_ORD_STATUS_515(6) = "���Њm�F�҂�"
    V_ORD_STATUS_515(14) = "����������"
    V_ORD_STATUS_515(8) = "��Ƃ�NG�A����"
    V_ORD_STATUS_515(7) = "���E�҂�NG�A����"
    V_ORD_STATUS_515(17) = "NG�I��"  
    V_ORD_STATUS_515(9) = "���ޏI��"
    V_ORD_STATUS_515(10) = "����"
'---- 2016.08.24 YenHTP.End ----
%>
<!--#include file="./5-15.htm"-->
<!--#include file="../include/sessionCn_Close.inc"-->
