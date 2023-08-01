<%
    Dim Conn, ConnStr
    ConnStr = "Provider=Sqloledb;Password=ximen12365abcD;Persist Security Info=True;User ID=sa;Initial Catalog=QPTreasureDB;Data Source=127.0.0.1, 1433;"
   	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.open ConnStr

    Dim no
    no = request.QueryString("no")
    Dim str
    str = "update QPTreasureDB.dbo.ChargeManage set is_delete = 1 where no=" & no
    Conn.Execute(str)

    Set rs = Nothing
    Conn.Close
    Set Conn = Nothing
    response.Redirect("payManage.asp")
%>