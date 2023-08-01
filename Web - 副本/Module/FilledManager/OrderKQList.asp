﻿<!--#include file="../../CommonFun.asp"-->
<!--#include file="../../GameConn.asp"-->
<!--#include file="../../Cls_Page.asp"-->
<!--#include file="../../conn.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
    <link href="../../Css/layout.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../../Js/common.js"></script>   
    <script type="text/javascript" src="../../Js/comm.js"></script>   
    <script type="text/javascript" src="../../Js/Check.js"></script>
    <script type="text/javascript" src="../../Js/Calendar.js"></script>
    
     <script type="text/javascript" src="../../Js/Sort.js"></script>  
</head>
<%
    If CxGame.GetRoleValue("205",trim(session("AdminName")))<"1" Then
    response.redirect "/Index.asp"   
    End If
%>
<body>
    <!-- 头部菜单 Start -->
    <table width="100%" border="0" cellpadding="0" cellspacing="0" class="title">
        <tr>
            <td width="19" height="25" valign="top"  class="Lpd10"><div class="arr"></div></td>
            <td width="1232" height="25" valign="top" align="left">你当前位置：充值系统 - 快钱管理</td>
        </tr>
    </table>
    <!-- 头部菜单 End -->
    <script type="text/javascript">        
        function QueryInfo(opType)
        {
            var startDate = document.form1.StartDate.value;
            var endDate = document.form1.EndDate.value;
            if(opType==1)
            {
                if(!isDate(startDate)||!isDate(endDate))
                {
                    alert("日期格式不正确！");
                    return
                }
            }
            if(compareDate(startDate,endDate)==false)
		    {
		        
		        alert("开始时间不可大于结束时间！")
		        return
		    }
            document.form1.QueryType.value = opType
            document.form1.submit();
        }
    </script>
    <form name="form1" action="" method="get">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="titleQueBg Tmg7">
            <tr>
                <td class="listTdLeft" style="width: 80px">
                    日期查询：
                </td>
                <td>
                    <% If CxGame.GetInfo(0,"StartDate")<>"" Then %>
                    <input name="StartDate" type="text" class="text" style="width:70px;" value="<%=CxGame.GetInfo(0,"StartDate") %>" maxlength="10" /><img src="../../Images/btn_calendar.gif" onclick="calendar(StartDate)" style="cursor:pointer;vertical-align:middle"/>                    
                    <% Else %>
                    <input name="StartDate" type="text" class="text" style="width:70px;" value="<%=Year(Date())&"-"&Month(Date())&"-1" %>" maxlength="10" /><img src="../../Images/btn_calendar.gif" onclick="calendar(StartDate)" style="cursor:pointer;vertical-align:middle"/>  
                    <% End If %>
                    至
                    <% If CxGame.GetInfo(0,"EndDate")<>"" Then %>
                    <input name="EndDate" type="text" class="text" style="width:70px;" value="<%=CxGame.GetInfo(0,"EndDate") %>" maxlength="10" /><img src="../../Images/btn_calendar.gif" onclick="calendar(EndDate)" style="cursor:pointer;vertical-align:middle"/>                    
                    <% Else%>
                    <input name="EndDate" type="text" class="text" style="width:70px;" value="<%=Date() %>"  maxlength="10"/><img src="../../Images/btn_calendar.gif" onclick="calendar(EndDate)" style="cursor:pointer;vertical-align:middle"/>                    
                    <% End If %>                    
                    <select name="payResult" style="width:80px;"> 
                    <option value="0">全部</option>
                    <option value="10"<%If Request("payResult")=10 Then %> selected="selected" <%End If %>>支付成功</option>
                    <option value="11"<%If Request("payResult")=11 Then %> selected="selected" <%End If %>>支出失败</option>
                    </select>           
                    <input type="button" value="查询" class="btn wd1" onclick="QueryInfo('1')" />
                    <input type="button" value="今天" class="btn wd1" onclick="QueryInfo('2')"/>
                    <input type="button" value="昨天" class="btn wd1" onclick="QueryInfo('3')"/>
                    <input type="button" value="本星期" class="btn wd6" onclick="QueryInfo('4')"/>
                    <input type="button" value="上星期" class="btn wd6" onclick="QueryInfo('5')"/> 
                    <input type="hidden" name="QueryType" />                               
                </td>
            </tr>
        </table>        
    </form> 
    <form name="form2" action="">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="titleQueBg Tmg7">
            <tr>
                <td class="listTdLeft" style="width: 80px">
                    普通查询：
                </td>
                <td>
                    <input name="search" type="text" class="text"  value="<%=CxGame.GetInfo(0,"search") %>"/>
                    <select name="swhat">
                    <option value="1" selected="selected">按订单号码</option>   
                    <option value="2"<% If CxGame.GetInfo(1,"swhat")=2 Then %> selected="selected" <%End If %>>按交易号码</option> 
                    <option value="3"<% If CxGame.GetInfo(1,"swhat")=3 Then %> selected="selected" <%End If %>>按支付地址</option> 
                    </select>
                    <input type="submit" value="查询" class="btn wd1" />                
                </td>
            </tr>
        </table>      
    </form>
    <% 
        Call ConnectGame(QPTreasureDB)
        Select case lcase(Request("action"))            
            case else
            Call Main()
        End Select
        Call CloseGame()  
        
        '删除操作
        Sub Delete(lID)
            Dim ID
            ID = Trim(lID)
            GameConn.execute "delete ReturnKQDetailInfo where DetailID='"&ID&"'"                
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
                    OrderStr=Replace(Request.QueryString("orders")," ","+")&" Asc "
                End If
                If Request.QueryString("OrderType")<>"-101" Then
                    OrderStr=Replace(Request.QueryString("orders")," ","+")&" Desc "
                End If
            Else 
                OrderStr = "DealTime DESC"
            End If
            queryCondition = " DetailID=0"
            Dim field
            field = "DetailID,OrderID,DealID,OrderAmount,PayAmount,Fee,OrderTime,DealTime,PayResult"
            '查询条件            
            If CxGame.GetInfo(0,"search")<>"" Then
                Select case Request("swhat")
                    case 1
                        queryCondition = " OrderID = '"&CxGame.GetInfo(0,"search")&"'"
                    case 2
                        queryCondition = " DealID = '"&CxGame.GetInfo(0,"search")&"'"
                    case 3
                        queryCondition = " Ext1 = '"&CxGame.GetInfo(0,"search")&"'"
                End Select
            End If    
            
            Dim startDate,endDate
            If CxGame.GetInfo(0,"QueryType")<>"" Then
                Select case CxGame.GetInfo(0,"QueryType")
                    case 1
                        startDate = CxGame.GetInfo(0,"StartDate")&" 00:00:00"
                        endDate = CxGame.GetInfo(0,"EndDate")&" 23:59:59"
                    case 2
                        startDate = Date()&" 00:00:00"
                        endDate = Date()&" 23:59:59"
                    case 3
                        startDate = DateAdd("d",-1,Date())&" 00:00:00"
                        endDate = DateAdd("d",-1,Date())&" 23:59:59"
                    case 4
                        startDate = DateAdd("d",1-Weekday(Date()),Date())&" 00:00:00"
                        endDate = DateAdd("d",6,startDate)&" 23:59:59"
                    case 5
                        startDate = DateAdd("d",1-7-Weekday(Date()),Date())&" 00:00:00"
                        endDate = DateAdd("d",6,startDate)&" 23:59:59"
                End Select
                queryCondition = "(DealTime between '"&startDate&"' And '"&endDate&"')"
            End If
            
            Select Case Request("payResult") 
                case 0
                    queryCondition = queryCondition&""
                case else
                    queryCondition = queryCondition&" And PayResult='"&Request("payResult")&"'"  
            End Select           
            
            '==============================================================================================================
            '执行存储过程
                    
            Set Page = new Cls_Page      '创建对象
            Set Page.Conn = GameConn     '得到数据库连接对象        
            With Page
                .PageSize = GetPageSize					'每页记录条数
                .PageParm = "page"				'页参数
                '.PageIndex = 10				'当前页，可选参数，一般是生成静态时需要
	            .Database = "mssql"				'数据库类型,AC为access,MSSQL为sqlserver2000存储过程版,MYSQL为mysql,PGSQL为PostGreSql
	            .Pkey="DetailID"					'主键
	            .Field=field	'字段
	            .Table="ReturnKQDetailInfo"			'表名
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
    <form name="myformList" method="post" action='?action=DeleteAll' onsubmit="return CheckFormList();">
        <% If Request("action") = "DeleteAll" Then %>
        <script type="text/javascript">
            showInfo("删除成功")
        </script>
        <% End If %>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="Tmg7">
            <tr>
                <td height="39" class="titleOpBg">
	                <input type="submit" value="删除" class="btn wd1" />                  
                </td>
            </tr>
        </table>  
        <div id="content">
            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="box">
                <tr align="center" class="bold">
                    <th class="listTitle1" width='38' align='center'><input type="checkbox" name="chkAll" onclick="SelectAll(this.checked);" /></th>
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('OrderID',this)">订单号码</a></th>
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('DealID',this)">快钱交易号</a></th>   
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('OrderAmount',this)">订单金额</a></th>                 
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('PayAmount',this)">实付金额</a></th>                    
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('Fee',this)">手续费</a></th>                    
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('OrderTime',this)">订单日期</a></th>
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('DealTime',this)">交易日期</a></th>
                    <th class="listTitle"><a class="l1"  href="" onclick="GetOrder('PayResult',this)">支付结果</a></th>                    
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
                %>
                <tr class="<%=className %>" onmouseover="currentcolor=this.style.backgroundColor;this.style.backgroundColor='#caebfc';this.style.cursor='default';"
                    onmouseout="this.style.backgroundColor=currentcolor">
                    <td>
                        <% If rs(8,i)=11 Then %>
                        <input name='cid' type='checkbox' value='<%=rs(0,i)%>'/>
                        <% End If %>
                    </td>    
                    <td><a class="l" href="javascript:void(0)" onclick="javascript:openWindowOwn('OrderKQInfo.asp?id=<%=rs(0,i) %>','<%=rs(1,i) %>',850,600)"><%=rs(1,i) %></a></td>
                    <td><%=rs(2,i) %></td>
                    <td><%=FormatNumber(rs(3,i),2,-1) %></td>
                    <td><%=FormatNumber(rs(4,i),2,-1) %></td>
                    <td><%=FormatNumber(rs(5,i),2,-1) %></td>
                    <td><%=rs(6,i) %></td>
                    <td><%=rs(7,i) %></td>
                    <td>
                        <%If rs(8,i) = "10" Then %>
                        <span style="color:blue">成功</span>
                        <%Else %>
                        <span style="color:red">失败</span>
                        <%End If %>
                    </td>
                </tr>
                <% 
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
	                <input type="submit" value="删除" class="btn wd1" />                  
                </td>
            </tr>
        </table>
    </form>
    <% 
        Set rs = Nothing
        End Sub
    %>
</body>
</html>
