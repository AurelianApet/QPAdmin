<%
    Dim type1, no
    type1 = request.QueryString("type")
    no = request.QueryString("no")
    Dim Conn, ConnStr
    ConnStr = "Provider=Sqloledb;Password=ximen12365abcD;Persist Security Info=True;User ID=sa;Initial Catalog=QPTreasureDB;Data Source=127.0.0.1, 1433;"
   	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.open ConnStr
    Dim str
    str = "update QPTreasureDB.dbo.ChargeManage set accept=" & type1 & " where no=" & no
    Dim amounts, accounts, golds, golds1;
'    response.Write(str)
    Conn.Execute(str)
    str = "select * from QPTreasureDB.dbo.ChargeManage where no=" & no
    Set rs = Conn.Execute(str)
    amounts = rs.Fields("money_amount")
    accounts = rs.Fields("useraccount")
    str = "select * from QPTreasureDB.dbo.ShareDetailInfo where Accounts=" & accounts
    Set rs = Conn.Execute(str)
    Dim amounts1
    amounts1 = rs.Fields("CardPrice") + amounts
    golds1 = rs.Fields("CardGold") + amounts * 100
    str = "update QPTreasureDB.dbo.ShareDetailInfo set CardPrice=" & amounts1 & " where Accounts=" & accounts
    Conn.Execute(str)
    
    str = "update QPTreasureDB.dbo.ShareDetailInfo set CardGold=" & golds1 & " where Accounts=" & accounts
    Conn.Execute(str)
'    rs.Close
    Set rs = Nothing
    Conn.Close
    Set Conn = Nothing
    response.Redirect("payManage.asp")
%>
<html>
    <head>
    </head>
<title>Processing...</title>
<body>
</body>
</html>