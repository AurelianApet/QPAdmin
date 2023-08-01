<!--#include file="../../CommonFun.asp"-->
<!--#include file="../../function.asp"-->
<!--#include file="../../GameConn.asp"-->
<!--#include file="../../Cls_Page.asp"-->
<!--#include file="../../conn.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>无标题页</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link href="../../Css/layout.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../../Js/common.js"></script>
    <script type="text/javascript" src="../../Js/comm.js"></script>
    <script type="text/javascript" src="../../Js/Check.js"></script>
    <script type="text/javascript" src="../../Js/Calendar.js"></script>
     <script type="text/javascript" src="../../Js/Sort.js"></script>
</head>
<%
    If CxGame.GetRoleValue("308",trim(session("AdminName")))<"1" Then
    response.redirect "/Index.asp"   
    End If
%>
<body>
    <!-- 头部菜单 Start -->
    <table width="100%" border="0" cellpadding="0" cellspacing="0" class="title">
        <tr>
            <td width="19" height="25" valign="top"  class="Lpd10"><div class="arr"></div></td>
            <td width="1232" height="25" valign="top" align="left">你当前位置：游戏用户 - 泡点设置</td>
        </tr>
    </table>
    <!-- 头部菜单 End -->
    <% 
        Call ConnectGame("QPPlatformDB")
        Select case lcase(Request("action"))    
            case "presentinfo" 
            Call GetPresentInfo()
            case "presentmax" 
            Call GetPresentMax()
            case "savemax"
            Call SaveMax()
            case "save"
            Call SavePresent()       
            case else
            Call Main()
        End Select
        Call CloseGame()  
        
        Sub SavePresent()
            Dim rs,sql
            Dim PresentMember
            Set rs=Server.CreateObject("Adodb.RecordSet")
            sql = "select * from GlobalPlayPresent where ServerID='"&Request("id")&"'"
            rs.Open sql,GameConn,1,3
            If rs.eof Then
                rs.addnew
            End If
            rs("ServerID") = CxGame.GetInfo(1,"in_ServerID")            
            PresentMember = CxGame.GetInfo(0,"in_PresentMember")
            rs("PresentMember") = Replace(PresentMember," ","")           
            rs("CellPlayPresnet") = CxGame.GetInfo(1,"in_CellPlayPresnet")
            rs("CellPlayTime") = CxGame.GetInfo(1,"in_CellPlayTime")   
            rs("StartPlayTime") = CxGame.GetInfo(1,"in_StartPlayTime")   
            rs("CellOnlinePresent") = CxGame.GetInfo(1,"in_CellOnlinePresent")
            rs("CellOnlineTime") = CxGame.GetInfo(1,"in_CellOnlineTime")      
            rs("StartOnlineTime") = CxGame.GetInfo(1,"in_StartOnlineTime")        
            rs("IsPlayPresent") = CxGame.GetInfo(1,"in_IsPlayPresent")            
            rs("IsOnlinePresent") = CxGame.GetInfo(1,"in_IsOnlinePresent")       
            rs.update
            If Request("id")<>"" Then
                Call CxGame.ShowInfo2("修改成功！","GrantList.asp?page="&Request("page")&"",1200)
            Else
                Call CxGame.ShowInfo2("新增成功！","GrantList.asp",1200)
            End If
            rs.close()
            Set rs = Nothing         
        End Sub
        
        Sub SaveMax()
            Dim rs,sql
            Set rs=Server.CreateObject("Adodb.RecordSet")
            sql = "select * from GlobalPlayPresent where ServerID=-3"
            rs.Open sql,GameConn,1,3
            If rs.eof Then
                rs.addnew
            End If
            rs("ServerID")=-3
            rs("MaxDatePresent") = CxGame.GetInfo(1,"in_MaxDatePresent")            
            rs("MaxPresent") = CxGame.GetInfo(1,"in_MaxPresent")            
            rs.update
            Call CxGame.ShowInfo2("封顶值设置成功！","GrantList.asp",1200)
            rs.close()
            Set rs = Nothing  
        End Sub
        
        '删除操作
        Sub Delete(lID)
            Dim ID
            ID = lID
            GameConn.execute "delete GlobalPlayPresent where ServerID='"&ID&"'"            
        End Sub
        
        Sub Main() 
            Dim cidArray, lLoop
            cidArray = Split(Request("cid"),",")
            For lLoop=0 To UBound(cidArray)
                Call Delete(cidArray(lLoop))
            Next           
            Dim rs,nav,Page,i
            Dim lCount, queryCondition, OrderStr
             If Request.QueryString("orders")<>""And Request.QueryString("OrderType")<>"" Then
                If Request.QueryString("OrderType")<>"-102" Then
                    OrderStr=Request.QueryString("orders")&" Asc "
                End If
                If Request.QueryString("OrderType")<>"-101" Then
                    OrderStr=Request.QueryString("orders")&" Desc "
                End If
            Else 
                OrderStr = "ServerID Asc"
            End If
            
            Dim field
            field = "*"
            '查询条件           
           
            
            '==============================================================================================================
            '执行存储过程
                    
            Set Page = new Cls_Page      '创建对象
            Set Page.Conn = GameConn     '得到数据库连接对象        
            With Page
                .PageSize = GetPageSize					'每页记录条数
                .PageParm = "page"				'页参数
                '.PageIndex = 10				'当前页，可选参数，一般是生成静态时需要
	            .Database = "mssql"				'数据库类型,AC为access,MSSQL为sqlserver2000存储过程版,MYSQL为mysql,PGSQL为PostGreSql
	            .Pkey="ServerID"					'主键
	            .Field=field	'字段
	            .Table="GlobalPlayPresent"			'表名
	            .Condition=queryCondition		'条件,不需要where
	            .OrderBy=OrderStr				'排序,不需要order by,需要asc或者desc
	            .RecordCount = 0				'总记录数，可以外部赋值，0不保存（适合搜索），-1存为session，-2存为cookies，-3存为applacation

	            .NumericJump = 5                '数字上下页个数，可选参数，默认为3，负数为跳转个数，0为显示所有
	            .Template = "总记录数：{$RecordCount} 总页数：{$PageCount} 每页记录数：{$PageSize} 当前页数：{$PageIndex} {$FirstPage} {$PreviousPage} {$NumericPage} {$NextPage} {$LastPage} {$InputPage} {$SelectPage}" '整体模板，可选参数，有默认值
	            .FirstPage = "首页"             '可选参数，有默认值
	            .PreviousPage = "上一页"        '可选参数，有默认值
	            .NextPage = "下一页"            '可选参数，有默认值
	            .LastPage = "尾页"              '可选参数，有默认值
	            .NumericPage = " {$PageNum} "   '数字分页部分模板，可选参数，有默认值
            End With
            
            rs = Page.ResultSet()       '记录集
            lCount = Page.RowCount()    '可选，输出总记录数
            nav = Page.Nav()            '分页样式
            
            Page = Null
            Set Page = Nothing
            '==============================================================================================================
    %>
    <script type="text/javascript">
        function CheckFormList()
        {
            if(!confirm("确定要执行选定的操作吗？"))
            {
                return false;
            }
            else
            {
                var cid="";
                var cids = GelTags("input");                
                for(var i=0;i<cids.length;i++)
                {
                    if(cids[i].checked)
                    {
                        if(cids[i].name == "cid")
                            cid+=cids[i].value+",";
                    }
                }       
                if(cid=="")
                {
                    showError("未选中任何数据");
                    return false;
                }
            }
        }
    </script>
    <form name="myformList" method="post" action='' onsubmit="return CheckFormList();">
        <% If Request("Action") = "DeleteAll" Then %>
        <script type="text/javascript">
            showInfo("删除成功")
        </script>
        <% End If %>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td height="39" class="titleOpBg">
                    <input type="button" value="新建" class="btn wd1" onclick="Redirect('?action=presentinfo')" />
                    <input class="btnLine" type="button" />
                    <input type="submit" value="删除" class="btn wd1" />
                    <input type="hidden" name="Action" value="DeleteAll" />        
                    <input type="button" value="封顶设置" class="btn wd2" onclick="Redirect('?action=presentmax')" />
                </td>
            </tr>
        </table>
        <div id="content">
            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="box">
                <tr class="bold">
                    <th class="listTitle1" width='38' align='center'><input type="checkbox" name="chkAll" onclick="SelectAll(this.checked);" /></th>                   
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('ServerID',this)">房间名称</a></th>                    
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('PresentMember',this)">赠送对象</a></th>                   
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('CellPlayPresnet',this)">游戏泡分单元值</a></th>
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('CellPlayTime',this)">游戏泡分单元时间</a></th>
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('StartPlayTime',this)">游戏泡分启始时间</a></th>
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('CellOnlinePresent',this)">在线泡分单元值</a></th>
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('CellOnlineTime',this)">在线泡分单元时间</a></th>
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('StartOnlineTime',this)">在线泡分启始时间</a></th>
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('IsPlayPresent',this)">游戏泡分状态</a></th>
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('IsOnlinePresent',this)">在线泡分状态</a></th>
                    <th class="listTitle"><a class="l1"  href="" onclick="GetOrder('CollectDate',this)">收录时间</a></th>
                </tr>
                <% 
                    Dim className
                    If IsNull(rs) Then
                        Response.Write("<tr class='tdbg'><td colspan='100' align='center'><br>没有任何信息!<br><br></td></tr>")
                    Else
                    
                    For i=0 To Ubound(rs,2)
                    If i Mod 2 = 0 Then
                        className="list"        
                    Else
                        className="listBg"      
                    End If 
                    
                    If rs(0,i)<>-3 Then
                %>
                <tr align="center" class="<%=className %>" onmouseover="currentcolor=this.style.backgroundColor;this.style.backgroundColor='#caebfc';this.style.cursor='pointer';"
                    onmouseout="this.style.backgroundColor=currentcolor">
                    <td><input name='cid' type='checkbox' value='<%=rs(0,i)%>'/></td>  
                    <td onclick="Redirect('?action=presentinfo&id=<%=rs(0,i) %>&page=<%=Request("page") %>')">
                    <% If rs(0,i)="-2" Then %>
                    积分房间
                    <% ElseIf rs(0,i)="-1" Then %>
                    金币房间
                    <% Else %>
                    <%=CxGame.GetRoomNameByRoomID(Trim(CInt(rs(0,i)))) %>
                    <% End If %>
                    </td>                 
                    <td onclick="Redirect('?action=presentinfo&id=<%=rs(0,i) %>&page=<%=Request("page") %>')">
                    <% 
                        Dim strObj
                        strObj=""
                        If InStr(rs(1,i),"0")>0 Then
                            IF strObj<>"" Then
                                strObj=strObj&",普通用户"
                            Else
                                strObj="普通用户"
                            End If
                        End IF
                        If InStr(rs(1,i),"1")>0 Then
                            IF strObj<>"" Then
                                strObj=strObj&",蓝钻会员"
                            Else
                                strObj="蓝钻会员"
                            End If
                        End IF
                        If InStr(rs(1,i),"2")>0 Then
                            IF strObj<>"" Then
                                strObj=strObj&",黄钻会员"
                            Else
                                strObj="黄钻会员"
                            End If
                        End IF
                        If InStr(rs(1,i),"3")>0 Then
                            IF strObj<>"" Then
                                strObj=strObj&",白钻会员"
                            Else
                                GrantObjet="白钻会员"
                            End If
                        End IF
                        If InStr(rs(1,i),"4")>0 Then
                            IF strObj<>"" Then
                                strObj=strObj&",红钻会员"
                            Else
                                strObj="红钻会员"
                            End If
                        End IF
                    %>
                    <%=strObj %>
                    </td>       
                    <td onclick="Redirect('?action=presentinfo&id=<%=rs(0,i) %>&page=<%=Request("page") %>')"><%=rs(4,i) %></td>   
                    <td onclick="Redirect('?action=presentinfo&id=<%=rs(0,i) %>&page=<%=Request("page") %>')"><%=rs(5,i) %></td>
                    <td onclick="Redirect('?action=presentinfo&id=<%=rs(0,i) %>&page=<%=Request("page") %>')"><%=rs(6,i) %></td>
                    <td onclick="Redirect('?action=presentinfo&id=<%=rs(0,i) %>&page=<%=Request("page") %>')"><%=rs(7,i) %></td>
                    <td onclick="Redirect('?action=presentinfo&id=<%=rs(0,i) %>&page=<%=Request("page") %>')"><%=rs(8,i) %></td>
                    <td onclick="Redirect('?action=presentinfo&id=<%=rs(0,i) %>&page=<%=Request("page") %>')"><%=rs(9,i) %></td>
                    <td onclick="Redirect('?action=presentinfo&id=<%=rs(0,i) %>&page=<%=Request("page") %>')">
                        <% If rs(10,i)=1 Then %>
                        开启
                        <% Else %>
                        关闭
                        <% End If %>
                    </td>
                    <td onclick="Redirect('?action=presentinfo&id=<%=rs(0,i) %>&page=<%=Request("page") %>')">
                        <% If rs(11,i)=1 Then %>
                        开启
                        <% Else %>
                        关闭
                        <% End If %>
                    </td>      
                    <td onclick="Redirect('?action=presentinfo&id=<%=rs(0,i) %>&page=<%=Request("page") %>')"><%=rs(12,i) %></td>    
                </tr>
                <% 
                    End If
                    Next
                    End If
                %>
            </table>
        </div>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td class="listTitleBg"><span>选择：</span>&nbsp;<a class="l1" href="javascript:SelectAll(true);">全部</a>&nbsp;-&nbsp;<a class="l1" href="javascript:SelectAll(false);">无</a></td>
                <td class="page" align="right"><%Response.Write nav%></td>
            </tr>
        </table> 
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" id="OpList">
            <tr>
                <td height="39" class="titleOpBg">
                    <input type="button" value="新建" class="btn wd1" onclick="Redirect('?action=presentinfo')" />
                    <input class="btnLine" type="button" />
                    <input type="submit" value="删除" class="btn wd1" />
                    <input type="hidden" name="Action" value="DeleteAll" />      
                    <input type="button" value="封顶设置" class="btn wd2" onclick="Redirect('?action=presentmax')" />   
                </td>
            </tr>
        </table> 
    </form>
    <% 
        End Sub
        
        Sub GetPresentInfo()
            Dim rs,sql,titleInfo
            Dim ServerID,PresentMember,CellPlayPresnet,CellPlayTime,StartPlayTime,CellOnlinePresent,CellOnlineTime,StartOnlineTime,IsPlayPresent,IsOnlinePresent,CollectDate
            Set rs=Server.CreateObject("Adodb.RecordSet")
            sql = "select * from GlobalPlayPresent where ServerID='"&Request("id")&"'"
            rs.Open sql,GameConn,1,3
            If rs.Bof And rs.Eof Then
                titleInfo = "新增泡点规则"
                ServerID=0
                PresentMember = ""
                CellPlayPresnet = 0
                CellPlayTime = 0
                StartPlayTime = 0
                CellOnlinePresent = 0
                CellOnlineTime = 0
                StartOnlineTime = 0
                IsPlayPresent = 1
                IsOnlinePresent = 1
            Else
                titleInfo = "修改泡点规则"
                ServerID = rs("ServerID")
                PresentMember = rs("PresentMember")
                CellPlayPresnet = rs("CellPlayPresnet")
                CellPlayTime = rs("CellPlayTime")
                StartPlayTime = rs("StartPlayTime")
                CellOnlinePresent = rs("CellOnlinePresent")
                CellOnlineTime = rs("CellOnlineTime")
                StartOnlineTime = rs("StartOnlineTime")
                IsPlayPresent = rs("IsPlayPresent")
                IsOnlinePresent = rs("IsOnlinePresent")
                CollectDate = rs("CollectDate")
            End If            
    %>
    <script type="text/javascript">
        function CheckFormInfo()
        {
            var CellPlayPresnet = document.myFormInfo.in_CellPlayPresnet.value;
            var CellPlayTime = document.myFormInfo.in_CellPlayTime.value;
            var StartPlayTime = document.myFormInfo.in_StartPlayTime.value;
            var CellOnlinePresent = document.myFormInfo.in_CellOnlinePresent.value;
            var CellOnlineTime = document.myFormInfo.in_CellOnlineTime.value;
            var StartOnlineTime = document.myFormInfo.in_StartOnlineTime.value;
            if(!IsPositiveInt(CellPlayPresnet))
            {
                alert("游戏泡分单元值必须为数字！");
                return false;
            }
            if(!IsPositiveInt(CellPlayTime))
            {
                alert("游戏泡分单元时间必须为数字！");
                return false;
            }
            if(!IsPositiveInt(StartPlayTime))
            {
                alert("游戏泡分启始时间必须为数字！");
                return false;
            }
            if(!IsPositiveInt(CellOnlinePresent))
            {
                alert("在线泡分单元值必须为数字！");
                return false;
            }
            if(!IsPositiveInt(CellOnlineTime))
            {
                alert("在线泡分单元时间必须为数字！");
                return false;
            }
            if(!IsPositiveInt(StartOnlineTime))
            {
                alert("在线泡分启始时间必须为数字！");
                return false;
            }
            return true;
        }
    </script>
    <form name="myFormInfo" method="post" action='?action=save&id=<%=Request("id") %>&page=<%=Request("page") %>' onsubmit="return CheckFormInfo()">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td class="titleOpBg Lpd10">
                    <input type="button" value="返回" class="btn wd1" onclick="Redirect('GrantList.asp')" />                                       
                    <input class="btnLine" type="button" />  
                    <input type="submit" value="保存" class="btn wd1" />  
                </td>
            </tr>
        </table>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="listBg2">
            <tr>
                <td height="35" colspan="2" class="f14 bold Lpd10 Rpd10"><div class="hg3  pd7"><%=titleInfo %></div></td>
            </tr>
            <tr>
                <td class="listTdLeft">房间名称：</td>
                <td>
                    <select name="in_ServerID"  style="width:120px;">
                    <option value="-2"<% If ServerID= "-2" Then %> selected="selected" <% End If %> >积分房间</option>
                    <option value="-1"<% If ServerID= "-1" Then %> selected="selected" <% End If %> >金币房间</option>
                    <% 
                        Dim ArrayKind,i
                        ArrayKind = CxGame.GetRoomList1()
                        For i=0 To Ubound(ArrayKind)                                                         
                    %>
                    <option value="<%=ArrayKind(i,0) %>"<% If ServerID= ArrayKind(i,0) Then %> selected="selected" <% End If %> ><%=ArrayKind(i,1) %></option>
                    <%                      
                        Next  
                        Set ArrayKind = nothing
                    %>
                    </select>          
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">赠送对象：</td>
                <td>
                    <input name="in_PresentMember" id="in_PresentMember1" type="checkbox" value="0"<% If InStr(PresentMember,"0")>0 Then %> checked="checked"<% End If %> /><label for="in_PresentMember1">普通用户</label>
                    <input name="in_PresentMember" id="in_PresentMember2" type="checkbox" value="1"<% If InStr(PresentMember,"1")>0 Then %> checked="checked"<% End If %> /><label for="in_PresentMember2">蓝钻会员</label> 
                    <input name="in_PresentMember" id="in_PresentMember3" type="checkbox" value="2"<% If InStr(PresentMember,"2")>0 Then %> checked="checked"<% End If %> /><label for="in_PresentMember3">黄钻会员</label> 
                    <input name="in_PresentMember" id="in_PresentMember4" type="checkbox" value="3"<% If InStr(PresentMember,"3")>0 Then %> checked="checked"<% End If %> /><label for="in_PresentMember4">白钻会员</label> 
                    <input name="in_PresentMember" id="in_PresentMember5" type="checkbox" value="4"<% If InStr(PresentMember,"4")>0 Then %> checked="checked"<% End If %> /><label for="in_PresentMember5">红钻会员</label>                                        
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">游戏泡分单元值：</td>
                <td>
                    <input name="in_CellPlayPresnet" type="text" class="text" value="<%=CellPlayPresnet %>" />                                    
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">游戏泡分单元时间（秒）：</td>
                <td>
                    <input name="in_CellPlayTime" type="text" class="text" value="<%=CellPlayTime %>" />                                   
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">游戏泡分启始时间（秒）：</td>
                <td>
                    <input name="in_StartPlayTime" type="text" class="text" value="<%=StartPlayTime %>" />                                   
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">在线泡分单元值：</td>
                <td>
                    <input name="in_CellOnlinePresent" type="text" class="text" value="<%=CellOnlinePresent %>" />                                    
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">在线泡分单元时间（秒）：</td>
                <td>
                    <input name="in_CellOnlineTime" type="text" class="text" value="<%=CellOnlineTime %>" />                                   
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">在线泡分启始时间（秒）：</td>
                <td>
                    <input name="in_StartOnlineTime" type="text" class="text" value="<%=StartOnlineTime %>" />                                   
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">泡分状态：</td>
                <td>
                    <input name="in_IsPlayPresent" id="in_IsPlayPresent" type="checkbox" value="1"<% If IsPlayPresent=1 Then %> checked="checked"<% End If %> /><label for="in_IsPlayPresent">开启游戏泡分</label>
                    <input name="in_IsOnlinePresent" id="in_IsOnlinePresent" type="checkbox" value="1"<% If IsOnlinePresent=1 Then %> checked="checked"<% End If %> /><label for="in_IsOnlinePresent">开启在线泡分</label>                        
                </td>
            </tr>
            <% If ServerID<>0 Then %>
            <tr>
                <td class="listTdLeft">创建时间：</td>
                <td>
                    <%=CollectDate %>                   
                </td>
            </tr>
            <% End If %>
        </table>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td class="titleOpBg Lpd10">
                    <input type="button" value="返回" class="btn wd1" onclick="Redirect('GrantList.asp')" />                                       
                    <input class="btnLine" type="button" />  
                    <input type="submit" value="保存" class="btn wd1" />  
                </td>
            </tr>
        </table>
    </form>
    <%
        rs.Close()
        Set rs=nothing
        End Sub
        
        Sub GetPresentMax()
            Dim rs,sql
            Dim ServerID,MaxDatePresent,MaxPresent
            Set rs=Server.CreateObject("Adodb.RecordSet")
            sql = "select * from GlobalPlayPresent where ServerID=-3"
            rs.Open sql,GameConn,1,3
            If rs.Bof And rs.Eof Then
                ServerID=-3
                MaxDatePresent = 0
                MaxPresent = 0
            Else
                ServerID = rs("ServerID")
                MaxDatePresent = rs("MaxDatePresent")
                MaxPresent = rs("MaxPresent")                
            End If            
    %>
    <script type="text/javascript">
        function CheckFormInfo2()
        {
            var MaxDatePresent = document.myFormInfo2.in_MaxDatePresent.value;
            var MaxPresent = document.myFormInfo2.in_MaxPresent.value;
            if(!IsPositiveInt(MaxDatePresent))
            {
                alert("单日封顶值必须为数字！");
                return false;
            }
            if(!IsPositiveInt(MaxPresent))
            {
                alert("总封顶值必须为数字！");
                return false;
            }            
            return true;
        }
    </script>
    <form name="myFormInfo2" method="post" action='?action=savemax&id=<%=Request("id") %>&page=<%=Request("page") %>' onsubmit="return CheckFormInfo2()">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td class="titleOpBg Lpd10">
                    <input type="button" value="返回" class="btn wd1" onclick="Redirect('GrantList.asp')" />                                       
                    <input class="btnLine" type="button" />  
                    <input type="submit" value="保存" class="btn wd1" />  
                </td>
            </tr>
        </table>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="listBg2">
            <tr>
                <td height="35" colspan="2" class="f14 bold Lpd10 Rpd10"><div class="hg3  pd7">泡点封顶设置(设置为0即为取消封顶)</div></td>
            </tr>
            <tr>
                <td class="listTdLeft">单日封顶值：</td>
                <td>
                    <input name="in_MaxDatePresent" type="text" class="text" value="<%=MaxDatePresent %>" />                                 
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">总封顶值：</td>
                <td>
                    <input name="in_MaxPresent" type="text" class="text" value="<%=MaxPresent %>" />                                 
                </td>
            </tr>            
        </table>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td class="titleOpBg Lpd10">
                    <input type="button" value="返回" class="btn wd1" onclick="Redirect('GrantList.asp')" />                                       
                    <input class="btnLine" type="button" />  
                    <input type="submit" value="保存" class="btn wd1" />  
                </td>
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