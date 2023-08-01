<!--#include file="../../CommonFun.asp"-->
<!--#include file="../../function.asp"-->
<!--#include file="../../GameConn.asp"-->
<!--#include file="../../conn.asp"-->
<!--#include file="../../md5.asp"-->
<!--#include file="../../Cls_Page.asp"-->
<%
dim aa
aa="000001"
'response.Write right("0000001",6)
'WriteTxtFile "0000001","jiqiren.txt"
'response.Write readTxtFile()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title></title>
    <link href="../../Css/layout.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        .querybox {width:500px; background:#caebfc;font-size:12px;line-height: 18px; text-align:left;
        border: 1px solid #066ba4;z-index:999; display:none;position: absolute; top:150px; left:200px;padding:5px;
        filter:progid:DXImageTransform.Microsoft.DropShadow(color=#9a8559,offX=1,offY=1,positives=true); }
    .STYLE1 {color: #FFFFFF}
    </style>
</head>
<%

    If CxGame.GetRoleValue("100",trim(session("AdminName")))<"1" Then
    response.redirect "/Index.asp"
    End If
	
	'写入记事本：写入最后批量添加帐号到记事本，以便下次添加从此处开始 RegAccounts
	Function WriteTxtFile(Text,FileName)
		path=Server.MapPath(FileName)
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f1 = fso.CreateTextFile(path,true)
		f1.Write (Text)
		f1.Close 
	End Function
	'读取记事本：读取上次最后添加的帐号，以便从这里开始重新添加 RegAccounts
	function readTxtFile()
		whichfile=server.mappath("jiqiren.txt")
		Set fso = CreateObject("Scripting.FileSystemObject")
		 Set txt = fso.OpenTextFile(whichfile,1)
		 readTxtFile=txt.ReadLine
	end function
  	'do until txt.AtEndOfStream
    ' rline =txt.ReadLine
  	'response.write ""&rline & "<br><br>"
 	 'loop

%>
<body>
    <%
if request("action")="add" and request("shuliang")<>"" then
	'变量gameID,NickNamev,FaceID,Gender
	dim newFaceid,newnub
	Call ConnectGame(QPAccountsDB)
	
	dim rs_j,togold
	set rs_j=GameConn.execute("select StatusName,StatusValue from dbo.SystemStatusInfo where StatusName='GrantScoreCount'")
	if not rs_j.eof then
		togold=rs_j("StatusValue")
	else
		response.Write("SystemStatusInfo")
		response.End
	end if
	
	newnub=0
	shuliang=request("shuliang") '
	for i=1 to shuliang
		dim endUserid,rs2,endGameID,newGameID
		set rs=GameConn.execute("select userid,GameID from dbo.AccountsInfo order by userid desc")
		if not rs.eof then
			endUserid=rs("userid")
			endGameID=rs("GameID")
		end if

		'-------
		dim a1,a2,newRegGameID
		if readTxtFile="" or readTxtFile=1 then
			a1="000001"
		else
			a1=readTxtFile()
		end if
			a2=a1+1
		newRegGameID=right("000000"&a2&"",6)
		'response.Write newRegGameID
		'response.End()
		'----------------
		
		newGameID=endGameID+1
		Accounts=newGameID
		NickName=newGameID
		RegAccounts=newRegGameID
		LogonPass="3FA0CA7E6984A471C6C14D0A4880DC9E"
		InsurePass="13B9314DE5BC831230EAE3311EAAF463"
		SpreaderID=0
		PassPortID=""
		Compellation=""
		if i mod 2 =0 then
			Gender=0	'
		else
			Gender=1
		end if
		
		Randomize timer
		newFaceid = cint(rnd*199+1)
		FaceID=newFaceid	'1-200
		GameLogonTimes=0
		LastLogonIP="127.0.0.1"
		LastLogonMachine="323B8D966D6060E768F0A15121C49103"
		RegisterIP="127.0.0.1"
		RegisterMachine="6B32B3463507A3EB2CE444381F545AD6"
		set rs2=GameConn.execute("insert into [AccountsInfo] (gameID,Accounts,NickName,RegAccounts,LogonPass,InsurePass,SpreaderID,PassPortID,Compellation,Gender,FaceID,GameLogonTimes,LastLogonIP,LastLogonMachine,RegisterIP,RegisterMachine) values('"&newGameID&"','"&Accounts&"','"&NickName&"','"&RegAccounts&"','"&LogonPass&"','"&InsurePass&"','"&SpreaderID&"','"&PassPortID&"','"&Compellation&"','"&Gender&"','"&FaceID&"','"&GameLogonTimes&"','"&LastLogonIP&"','"&LastLogonMachine&"','"&RegisterIP&"','"&RegisterMachine&"')")
		
		'写入到记事本
		WriteTxtFile newRegGameID,"jiqiren.txt"
		
		'userID
		dim rs_u,newuserid
		set rs_u=GameConn.execute("select userid,Accounts,LastLogonIP,RegisterIP from AccountsInfo where Accounts='"&Accounts&"'")
		newuserid=rs_u("userid")
		LastLogonIP=rs_u("LastLogonIP")
		RegisterIP=rs_u("RegisterIP")
		'插入记录
		dim rs_f,newDateID
		set rs_f=GameConn.execute("select DateID from [SystemGrantCount] order by DateID desc")
		newDateID=rs_f("DateID")+1
		set rs=GameConn.execute("insert into [SystemGrantCount] (DateID, RegisterIP, RegisterMachine, GrantScore, GrantCount) VALUES ('"&newDateID&"','"&RegisterIP&"','','"&togold&"','1')")
		'-------------------------------
		'写入金币
		'Call ConnectGame(QPTreasureDB)
		dim rs3
		set rs3=GameConn.execute("insert into QPTreasureDBLink.QPTreasureDB.dbo.GameScoreInfo (UserID, Score, RegisterIP, LastLogonIP) values('"&newuserid&"',"&togold&",'"&RegisterIP&"','"&LastLogonIP&"')")
		newnub=newnub+1
	next
		response.Write("成功添加"&newnub&"个帐号")
		response.end
end if
	%>
    <table width="100%" border="0" cellpadding="0" cellspacing="0" class="title">
        <tr>
            <td width="19" height="25" valign="top"  class="Lpd10"><div class="arr"></div></td>
            <td width="1232" height="25" valign="top" align="left">你当前位置：游戏用户 - 批量添加用户</td>
        </tr>
    </table>
    <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#E4E4E5">
	<form name="form" action="AdduserMore.asp?action=add" method="post">
      <tr>
        <td align="right" bgcolor="#006699">&nbsp;</td>
        <td bgcolor="#006699"><span class="STYLE1">性别、头像等都是自动随机添加</span></td>
      </tr>
      <tr>
        <td width="30%" align="right" bgcolor="#FFFFFF">帐号数量：</td>
        <td width="70%" bgcolor="#FFFFFF"><label>
          <input name="shuliang" type="text" id="shuliang" size="8" />
        个</label></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF">&nbsp;</td>
        <td bgcolor="#FFFFFF"><label>
          <input type="submit" name="Submit" value="确定帐号" />
        </label></td>
      </tr></form>
</table>
    
</body>
</html>
