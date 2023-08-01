<!--#include file="../../CommonFun.asp"-->
<!--#include file="../../function.asp"-->
<!--#include file="../../GameConn.asp"-->
<!--#include file="../../conn.asp"-->
<!--#include file="../../md5.asp"-->
<!--#include file="../../Cls_Page.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
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
    If CxGame.GetRoleValue("305",trim(session("AdminName")))<"1" Then
    response.redirect "/Index.asp"   
    End If
%>

<body>

    <!-- 头部菜单 Start -->
    <table width="100%" border="0" cellpadding="0" cellspacing="0" class="title">
        <tr>
            <td width="19" height="25" valign="top"  class="Lpd10"><div class="arr"></div></td>
            <td width="1232" height="25" valign="top" align="left">目前操作功能：道具管理</td>
        </tr>
    </table>
    <!-- 头部菜单 End -->
<%
        Call ConnectGame(QPTreasureDB)
        Select case Request("action")
            case "flowercateinfo"
                Call GetFlowerCateInfo()
            case "save"
                Call SaveFlowerCateInfo()
            case else
                Call Main()
        End Select 
        Call CloseGame()
     
       '删除
        Sub Delete(lID)
            Dim ID,KindID,dname
            ID = lID
            GameConn.execute "Delete From GameProperty Where ID="&lID
        End Sub
        
        '禁用还原
        Sub Nullity(lID,typeValue)
            Dim ID,KindID,dname
            ID = lID
            GameConn.execute "Update GameProperty Set Nullity="&typeValue&" Where ID="&lID
        End Sub
        
       '更新
        Sub SaveFlowerCateInfo()
            Dim rs,sql
            Dim IssueAreaArr,IssueArea,i
            Dim ServiceAreaArr,ServiceArea,j
            Set rs=Server.CreateObject("Adodb.RecordSet")
            sql = "select * from GameProperty where ID='"&Request("id")&"'"
            rs.Open sql,GameConn,1,3
            If rs.eof Then
                rs.addnew
            End If
            rs("Name") = CxGame.GetInfo(0,"in_Name")
            rs("Cash") = CxGame.GetInfo(0,"in_Cash")
            rs("Gold") = CxGame.GetInfo(0,"in_Gold")
            rs("Discount") = CxGame.GetInfo(0,"in_Discount")        
            '发行范围
            IssueAreaArr = Split(CxGame.GetInfo(0,"in_IssueArea"),",")
            IssueArea=0       
            For i=0 To UBound(IssueAreaArr)
                IssueArea = IssueArea Or IssueAreaArr(i)
            Next
            rs("IssueArea") =IssueArea
            '使用范围
            ServiceAreaArr = Split(CxGame.GetInfo(0,"in_ServiceArea"),",")
            ServiceArea=0
            For j=0 To UBound(ServiceAreaArr)
                ServiceArea = ServiceArea Or ServiceAreaArr(j)
            Next              
            rs("ServiceArea") = ServiceArea
            
            rs("SendLoveLiness") = CxGame.GetInfo(0,"in_SendLoveLiness")            
            rs("RecvLoveLiness") = CxGame.GetInfo(0,"in_RecvLoveLiness")            
            rs("RegulationsInfo") = CxGame.GetInfo(0,"in_RegulationsInfo")            
            rs("Nullity") = CxGame.GetInfo(1,"in_Nullity")
            rs.update
            If Request("id")<>"" Then
                Call CxGame.ShowInfo2("修改成功！","GamePropertyManager.asp?page="&Request("page")&"",1200)
            Else
                Call CxGame.ShowInfo2("新增成功！","GamePropertyManager.asp",1200)
            End If
            rs.close()
            Set rs = Nothing    
        End Sub
       
     Sub Main()
            '操作
            Dim cidArray, lLoop,acType
            cidArray = Split(Request("cid"),",")
            acType = Request("action")
            For lLoop=0 To UBound(cidArray)
                Select case acType
                    case "del"
                        Call Delete(cidArray(lLoop))
                    case "jinyong"
                        Call Nullity(cidArray(lLoop),1)
                    case "huanyuan"
                        Call Nullity(cidArray(lLoop),0)
                End Select
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
                OrderStr = "ID ASC"
            End If
            Dim field
            field = "ID,Name,Cash,Gold,Discount,IssueArea,ServiceArea,SendLoveLiness,RecvLoveLiness,Nullity"
          
            Dim startDate,endDate
            
           
            Dim sql,userID
           
           
            '==============================================================================================================
            '执行存储过程
                    
            Set Page = new Cls_Page      '创建对象
            Set Page.Conn = GameConn     '得到数据库连接对象        
            With Page
                .PageSize = GetPageSize					'每页记录条数
                .PageParm = "page"				'页参数
                '.PageIndex = 10				'当前页，可选参数，一般是生成静态时需要
	            .Database = "mssql"				'数据库类型,AC为access,MSSQL为sqlserver2000存储过程版,MYSQL为mysql,PGSQL为PostGreSql
	            .Pkey=""					'主键
	            .Field=field	'字段
	            .Table="GameProperty"			'表名
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
        function Operate(opType)
        {
            var opVal = document.myformList.in_optype
            if(!confirm("确定要执行选定的操作吗？"))
            {
                return;
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
                    return;
                }
                //操作处理
                if(opType=="del")
                    opVal.value = "del";
                if(opType=="jinyong")
                    opVal.value = "jinyong";
                if(opType=="huanyuan")
                    opVal.value = "huanyuan";
            }
           
            document.myformList.action = "?action="+opVal.value;
            document.myformList.submit();
        }
          
      
    </script>
    <form name="myformList" method="post" action=''>
        <% If Request("action") = "del" Then %>
        <script type="text/javascript">
            showInfo("删除成功")
        </script>
        <% End If %>
        <% If Request("action") = "jinyong" Then %>
        <script type="text/javascript">
            showInfo("禁用成功")
        </script>
        <% End If %>
        <% If Request("action") = "huanyuan" Then %>
        <script type="text/javascript">
            showInfo("还原成功")
        </script>
        <% End If %>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" >
            <tr>
                <td height="39" class="titleOpBg">
                    <%
                        If trim(session("AdminName"))="admin" or CxGame.GetRoleValue("305",trim(session("AdminName")))="3" or CxGame.GetRoleValue("305",trim(session("AdminName")))="1" or CxGame.GetRoleValue("305",trim(session("AdminName")))="7" Then
                    %>
	                <input type="button" value="删除" class="btn wd1" onclick="Operate('del')" />
                    <input type="button" value="禁用" class="btn wd1" onclick="Operate('jinyong')"/>
                    <input type="button" value="还原" class="btn wd1" onclick="Operate('huanyuan')"/>
                     <input name="in_optype" type="hidden" />  
                    <%
                        End If
                        If trim(session("AdminName"))="admin" or CxGame.GetRoleValue("305",trim(session("AdminName")))>"3" Then
                    %>       
                    <%
                        End If
                    %>    
                </td>
            </tr>
        </table>  
        <div id="content">
            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="box">
                <tr align="center" class="bold">
                    <th class="listTitle1" width='38' align='center'><input type="checkbox" name="chkAll" onclick="SelectAll(this.checked);" /></th>                    
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('Name',this)">道具名称</a></th>           
                    <th class="listTitle2"><a class="l1"  href="" onclick="GetOrder('Cash',this)">现金价格</a></th>              
                    <th class="listTitle2"><a  class="l1" href="" onclick="GetOrder('Gold',this)">金币价格</a></th>
                    <th class="listTitle2"><a  class="l1" href="" onclick="GetOrder('Discount',this)">折扣价</a></th>   
                    <th class="listTitle2"><a  class="l1" href="" onclick="GetOrder('IssueArea',this)">发行范围</a></th>  
                    <th class="listTitle2"><a  class="l1" href="" onclick="GetOrder('ServiceArea',this)">使用范围</a></th>  
                    <th class="listTitle2"><a  class="l1" href="" onclick="GetOrder('SendLoveLiness',this)">赠送魅力</a></th>  
                    <th class="listTitle2"><a  class="l1" href="" onclick="GetOrder('RecvLoveLiness',this)">接收魅力</a></th>    
                    <th class="listTitle"><a  class="l1" href="" onclick="GetOrder('Nullity',this)">禁止标识</a></th>               
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
                 <tr class="<%=className %>" onmouseover="currentcolor=this.style.backgroundColor;this.style.backgroundColor='#caebfc';this.style.cursor='pointer';"
                    onmouseout="this.style.backgroundColor=currentcolor"> 
                    <td><input name='cid' type='checkbox' value='<%=rs(0,i)%>'/></td>                  
                    <td onclick="Redirect('GamePropertyManager.asp?action=flowercateinfo&id=<%=rs(0,i) %>')"><%=rs(1,i) %></td>   
                    <td onclick="Redirect('GamePropertyManager.asp?action=flowercateinfo&id=<%=rs(0,i) %>')"><%=rs(2,i) %></td>
                    <td onclick="Redirect('GamePropertyManager.asp?action=flowercateinfo&id=<%=rs(0,i) %>')"><%=rs(3,i) %></td>  
                    <td onclick="Redirect('GamePropertyManager.asp?action=flowercateinfo&id=<%=rs(0,i) %>')"><%=rs(4,i) %></td> 
                    <td onclick="Redirect('GamePropertyManager.asp?action=flowercateinfo&id=<%=rs(0,i) %>')"><%=CxGame.GetIssueAreae(rs(5,i)) %></td>
                    <td onclick="Redirect('GamePropertyManager.asp?action=flowercateinfo&id=<%=rs(0,i) %>')"><%=CxGame.GetServiceArea(rs(6,i)) %></td>  
                    <td onclick="Redirect('GamePropertyManager.asp?action=flowercateinfo&id=<%=rs(0,i) %>')"><%=rs(7,i) %></td>  
                    <td onclick="Redirect('GamePropertyManager.asp?action=flowercateinfo&id=<%=rs(0,i) %>')"><%=rs(8,i) %></td>    
                    <td onclick="Redirect('GamePropertyManager.asp?action=flowercateinfo&id=<%=rs(0,i) %>')"><%=CxGame.GetNullityInfo(rs(9,i)) %></td>                      
                </tr>
                <% 
                    Next                    
                    End If
                %>
            </table>           
            </div> 
            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                    <td class="listTitleBg" style="height: 19px"><span>选择：</span>&nbsp;<a class="l1" href="javascript:SelectAll(true);">全部</a>&nbsp;-&nbsp;<a class="l1" href="javascript:SelectAll(false);">无</a></td>
                    <td class="page" align="right"><%Response.Write nav%></td>
                </tr>
            </table>  
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" id="OpList">
            <tr>
                <td height="39" class="titleOpBg">
	                <%
                        If trim(session("AdminName"))="admin" or CxGame.GetRoleValue("305",trim(session("AdminName")))="3" or CxGame.GetRoleValue("305",trim(session("AdminName")))="1" or CxGame.GetRoleValue("305",trim(session("AdminName")))="7" Then
                    %>                     
	                <input type="button" value="删除" class="btn wd1" onclick="Operate('del')" />
                    <input type="button" value="禁用" class="btn wd1" onclick="Operate('jinyong')"/>
                    <input type="button" value="还原" class="btn wd1" onclick="Operate('huanyuan')"/>
                    <%
                        End If
                        If trim(session("AdminName"))="admin" or CxGame.GetRoleValue("305",trim(session("AdminName")))>"3" Then
                    %>       
                    <%
                        End If
                    %>  
                </td>
            </tr>
        </table>  
    </form>
    <% 
        End Sub
        
        Sub GetFlowerCateInfo()
            Dim rs,sql
            Dim id,Name,Cash,SendUserCharm,RcvUserCharm,DisCount,IssueArea,ServiceArea,RegulationsInfo,Gold,Nullity
            If Request("id")<>"" Then
                Set rs=Server.CreateObject("Adodb.RecordSet")
                Call ConnectGame(QPTreasureDB)
                sql="select *  from  GameProperty where  ID="&Request("id")
                rs.Open sql,GameConn,1,3
                id=rs("ID")
                Name=rs("Name")
                Cash=rs("Cash")
                Gold=rs("Gold")
                SendUserCharm=rs("SendLoveLiness")
                RcvUserCharm=rs("RecvLoveLiness")
                DisCount=rs("DisCount")
                IssueArea=rs("IssueArea")
                ServiceArea=rs("ServiceArea")
                RegulationsInfo=rs("RegulationsInfo")
                Nullity=rs("Nullity")
            End If
    %>
    <script type="text/javascript">
        function CheckFormInfo()
        {
            var name = document.form1.Name.value;
            var cash = document.form1.Cash.value;
            var gold = document.form1.Gold.value;
            var sendUserCharm = document.form1.SendUserCharm.value;
            var rcvUserCharm = document.form1.RcvUserCharm.value;
            var disCount = document.form1.DisCount.value;
            var regulationsInf = document.form1.RegulationsInfo.value;
            
            if(Trim(name)=="")
            {
                alert("道具名称不能为空！");
                return false;
            }
            else if(len(name)>32)
            {
                alert("道具名称字符长度不可超过32个字符！");
                return false;
            }
            if(IsPositiveInt(cash)==false)
            {
                alert("道具价格非数值型字符！");
                return false;
            }
            if(IsPositiveInt(gold)==false)
            {
                alert("道具金币非数值型字符！");
                return false;
            }
            if(IsPositiveInt(disCount)==false)
            {
                alert("折扣价非数值型字符！");
                return false;
            }
            if(IsPositiveInt(sendUserCharm)==false)
            {
                alert("赠送魅力值非数值型字符！");
                return false;
            }
            if(IsPositiveInt(rcvUserCharm)==false)
            {
                alert("接收魅力值非数值型字符！");
                return false;
            }
            if(Trim(regulationsInf)=="")
            {
                alert("道具说明不能为空！");
                return false;
            }
            else if(len(regulationsInf)>255)
            {
                alert("道具说明字符长度不可超过255个字符！")
                return false;
            }
        }
    </script>
     <form name="form1" method="post" action='?action=save&id=<%=id %>' onsubmit="return CheckFormInfo()">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td class="titleOpBg Lpd10">          
                    <input type="submit" value="保存" class="btn wd1" />   
                    <input type="button" value="返回" class="btn wd1" onclick="Redirect('GamePropertyManager.asp')" />      
                </td>
            </tr>
        </table>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="listBg2">
            <tr>
                <td height="35" colspan="2" class="f14 bold Lpd10 Rpd10"><div class="hg3  pd7">修改道具信息</div></td>
            </tr>
            <tr>
	            <td class="listTdLeft"> 道具名称：</td>
	            <td>
	                <input name="in_Name" type="text" class="text wd4" value="<%=Name %>"/>       
	            </td>
	        </tr>
	        <tr>
	            <td class="listTdLeft">现金价格：</td>
	            <td>
	                <input name="in_Cash" type="text" class="text wd4" value="<%=Cash %>"/>（网站使用）                  
	            </td>
	        </tr> 
	        <tr>
	            <td class="listTdLeft">金币价格：</td>
	            <td>
	                <input name="in_Gold" type="text" class="text wd4" value="<%=Gold %>"/>                  
	            </td>
	        </tr> 
	        <tr>
	            <td class="listTdLeft">折扣：</td>
	            <td>
	                <input name="in_Discount" type="text" class="text wd4" value="<%=DisCount %>" />%              
	            </td>
	        </tr>
	        <tr>
	            <td class="listTdLeft">赠送魅力值：</td>
	            <td>
	                <input name="in_SendLoveLiness" type="text" class="text wd4" value="<%=SendUserCharm %>" />
	            </td>
	        </tr>
	        <tr>
	            <td class="listTdLeft">接收魅力值：</td>
	            <td>
	                <input name="in_RecvLoveLiness" type="text" class="text wd4" value="<%=RcvUserCharm %>" />
	            </td>
	        </tr>
	        <tr>
	            <td class="listTdLeft">发行范围：</td>
	            
	            <td>
	                <input name="in_IssueArea" id="in_IssueArea1" type="checkbox" value="1"<% If IssueArea And 1 Then %> checked='checked'<% End If %>/><label for="in_IssueArea1">商城道具</label>        
                    <input name="in_IssueArea" id="in_IssueArea2" type="checkbox" value="2"<% If IssueArea And 2 Then %> checked='checked'<% End If %>/><label for="in_IssueArea2">游戏道具</label>
                    <input name="in_IssueArea" id="in_IssueArea3" type="checkbox" value="4"<% If IssueArea And 4 Then %> checked='checked'<% End If %>/><label for="in_IssueArea3">房间道具</label>
	            </td>
	        </tr>
	         <tr>
	            <td class="listTdLeft">使用范围：</td>
	            <td>
	                <input name="in_ServiceArea" id="in_ServiceArea1" type="checkbox" value="1"<% If ServiceArea And 1 Then %> checked='checked'<% End If %>/><label for="in_ServiceArea1">自己范围</label>        
                    <input name="in_ServiceArea" id="in_ServiceArea2" type="checkbox" value="2"<% If ServiceArea And 2 Then %> checked='checked'<% End If %>/><label for="in_ServiceArea2">玩家范围</label>
                    <input name="in_ServiceArea" id="in_ServiceArea3" type="checkbox" value="4"<% If ServiceArea And 4 Then %> checked='checked'<% End If %>/><label for="in_ServiceArea3">旁观范围</label>             
	            </td>
	        </tr>	        
	        <tr>
                <td class="listTdLeft">使用说明：</td>
                <td>
                    <input name="in_RegulationsInfo" type="text" class="text wd4" style="width:500px;" value="<%=RegulationsInfo %>"/>&nbsp;使用说明字符长度不可超过256个字符                                                           
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">禁用状态：</td>
                <td>
                    <input name="in_Nullity" id="in_Nullity1" type="radio" value="0"<% If Nullity = 0 Then %> checked="checked"<% End If  %> /><label for="in_Nullity1">启用</label> 
                    <input name="in_Nullity" id="in_Nullity2" type="radio" value="1"<% If Nullity = 1 Then %> checked="checked"<% End If  %> /><label for="in_Nullity2">冻结</label> 
                </td>
            </tr>
        </table>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td class="titleOpBg Lpd10">            
                    <input type="submit" value="保存" class="btn wd1" />   
                    <input type="button" value="返回" class="btn wd1" onclick="Redirect('GamePropertyManager.asp')" />    
                </td>
            </tr>
        </table>    
    </form>
    <%
        End Sub 
     %>
</body>
</html>
