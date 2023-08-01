<!--#include file="../../CommonFun.asp"-->
<!--#include file="../../GameConn.asp"-->
<!--#include file="../../conn.asp"-->
<!--#include file="../../MBCard.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title><%=GetQPAdminSiteName() %></title>
    <link href="../../Css/layout.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../../Js/common.js"></script>
</head>
<body>
    <!-- 头部菜单 Start -->
    <table width="100%" border="0" cellpadding="0" cellspacing="0" class="title">
        <tr>
            <td width="19" height="25" valign="top"  class="Lpd10"><div class="arr"></div></td>
            <td width="1232" height="25" valign="top" align="left">目前操作功能：用户密保卡修改</td>
        </tr>
    </table>
    <!-- 头部菜单 End -->
    <% 
        Call ConnectGame(QPAccountsDB)
        Select case lcase(Request("action"))
            case "del"
            Call Delete()
            case else
            Call Main()
        End Select
        Call CloseGame()
        
        '删除
        Sub Delete()
            Dim sql
            sql = "Update AccountsInfo Set PasswordID=0 where UserID="&Request("UserID")
            response.Write sql
            GameConn.execute sql
            Response.Write "<script>opener.document.location.href='AccountsInfo.asp?id="&Request("UserID")&"';window.close();</script>" 
        End Sub        
       
        Sub Main()
            Dim rs,sql
            Set rs=Server.CreateObject("Adodb.RecordSet")
            sql = "select * from AccountsInfo where UserID="&Request("UserID")
            rs.Open sql,GameConn,1,3 
    %>
    <form name="form1" method="post" action="?UserID=<%=Request("UserID") %>&action=del">
        <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td class="titleOpBg Lpd10"> 
                    <input type="button" value="关闭" class="btn wd1" onclick="window.close();" />
                    <input type="submit" value="取消密保" class="btn wd2" />                    
                    <input name="in_optype" type="hidden" />
                </td>
            </tr>
        </table>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="box Tmg7">
            <tr align="center" class="bold">
                <td align="left" style="padding-left:10px;padding-bottom:8px;" colspan="7">序列号：<%=GetNewPassWordID(rs("PasswordID")) %></td>
            </tr>
            <tr align="center" class="bold">
                <th class="listTitle1">&nbsp;</th>
                <th class="listTitle2">1</th>
                <th class="listTitle2">2</th>
                <th class="listTitle2">3</th>
                <th class="listTitle">4</th>   
            </tr> 
            <tr align="center" class="list">
                <td class="leftTop trBg bold" style="width:50px;">A</td>
                <td><%=GetPasswordNum(rs("PasswordID"),A1) %></td>
                <td><%=GetPasswordNum(rs("PasswordID"),A2) %></td>
                <td><%=GetPasswordNum(rs("PasswordID"),A3) %></td>
                <td><%=GetPasswordNum(rs("PasswordID"),A4) %></td>                        
            </tr>
            <tr align="center" class="listBg">
                <td class="leftTop trBg bold" style="width:50px;">B</td>
                <td><%=GetPasswordNum(rs("PasswordID"),B1) %></td>
                <td><%=GetPasswordNum(rs("PasswordID"),B2) %></td>
                <td><%=GetPasswordNum(rs("PasswordID"),B3) %></td>
                <td><%=GetPasswordNum(rs("PasswordID"),B4) %></td>                                
            </tr>
            <tr align="center" class="list">
                <td class="leftTop trBg bold" style="width:50px;">C</td>
                <td><%=GetPasswordNum(rs("PasswordID"),C1) %></td>
                <td><%=GetPasswordNum(rs("PasswordID"),C2) %></td>
                <td><%=GetPasswordNum(rs("PasswordID"),C3) %></td>
                <td><%=GetPasswordNum(rs("PasswordID"),C4) %></td>                             
            </tr>
            <tr align="center" class="listBg">
                <td class="leftTop trBg bold" style="width:50px;">D</td>
                <td><%=GetPasswordNum(rs("PasswordID"),D1) %></td>
                <td><%=GetPasswordNum(rs("PasswordID"),D2) %></td>
                <td><%=GetPasswordNum(rs("PasswordID"),D3) %></td>
                <td><%=GetPasswordNum(rs("PasswordID"),D4) %></td>                            
            </tr>
            <tr align="center" class="list">
                <td class="leftTop trBg bold" style="width:50px;">E</td>
                <td><%=GetPasswordNum(rs("PasswordID"),E1) %></td>
                <td><%=GetPasswordNum(rs("PasswordID"),E2) %></td>
                <td><%=GetPasswordNum(rs("PasswordID"),E3) %></td>
                <td><%=GetPasswordNum(rs("PasswordID"),E4) %></td>                              
            </tr>
            <tr align="center" class="listBg">
                <td class="leftTop trBg bold" style="width:50px;">F</td>
                <td><%=GetPasswordNum(rs("PasswordID"),F1) %></td>
                <td><%=GetPasswordNum(rs("PasswordID"),F2) %></td>
                <td><%=GetPasswordNum(rs("PasswordID"),F3) %></td>
                <td><%=GetPasswordNum(rs("PasswordID"),F4) %></td>                          
            </tr>
        </table>
    </form>
    <% 
        rs.Close()
        Set rs=nothing
        End Sub
    %>
</body>
</html>
