<%
    Dim Conn, ConnStr
    ConnStr = "Provider=Sqloledb;Password=ximen12365abcD;Persist Security Info=True;User ID=sa;Initial Catalog=QPTreasureDB;Data Source=127.0.0.1, 1433;"
   	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.open ConnStr

    Dim no, recharege_no, discharge_no, total_recharge_no, total_discharge_no, recharge_wait, discharge_wait, recharge_accept, discharge_accept, recharge_cancel, discharge_cancel
    Set rs = Conn.Execute("select MAX(no) from QPTreasureDB.dbo.ChargeManage")
    if Not rs.EOF then
        no = rs.Fields(0) + 1
    end if
    Set rs = Conn.Execute("select count(*) from QPTreasureDB.dbo.ChargeManage where recharge_field=1 and accept=0")
    if Not rs.EOF then
        recharge_no = rs.Fields(0)
    end if
    Set rs = Conn.Execute("select count(*) from QPTreasureDB.dbo.ChargeManage where discharge_field=1 and accept=0")
    if Not rs.EOF then
        discharge_no = rs.Fields(0)
    end if
    Set rs = Conn.Execute("select count(*) from QPTreasureDB.dbo.ChargeManage where recharge_field=1")
    if Not rs.EOF then
        total_recharge_no = rs.Fields(0)
    end if
    Set rs = Conn.Execute("select count(*) from QPTreasureDB.dbo.ChargeManage where discharge_field=1")
    if Not rs.EOF then
        total_discharge_no = rs.Fields(0)
    end if

    Set rs = Conn.Execute("select count(*) from QPTreasureDB.dbo.ChargeManage where recharge_field=1 and accept=1")
    if rs.EOF = false then
        recharge_accept = rs.Fields(0)
    end if
    Set rs = Conn.Execute("select count(*) from QPTreasureDB.dbo.ChargeManage where recharge_field=1 and accept=2")
    if Not rs.EOF then
        recharge_wait = rs.Fields(0)
    end if
    Set rs = Conn.Execute("select count(*) from QPTreasureDB.dbo.ChargeManage where recharge_field=1 and accept=3")
    if Not rs.EOF then
        recharge_cancel = rs.Fields(0)
    end if
    Set rs = Conn.Execute("select count(*) from QPTreasureDB.dbo.ChargeManage where discharge_field=1 and accept=1")
    if Not rs.EOF then
        discharge_accept = rs.Fields(0)
    end if
    Set rs = Conn.Execute("select count(*) from QPTreasureDB.dbo.ChargeManage where discharge_field=1 and accept=2")
    if Not rs.EOF then
        discharge_wait = rs.Fields(0)
    end if
    Set rs = Conn.Execute("select count(*) from QPTreasureDB.dbo.ChargeManage where discharge_field=1 and accept=3")
    if Not rs.EOF then
        discharge_cancel = rs.Fields(0)
    end if
    Set rs = Conn.Execute("select * from QPTreasureDB.dbo.ChargeManage where accept=0")

'    Dim rs1
'    While Not rs.EOF
'        rs1 = Conn.Execute("select * from QPTreasureDB.dbo.standby where id=" & rs.Fields("no"))
'        response.Write("select * from QPTreasureDB.dbo.standby where id=" & rs.Fields("no"))
'        if Not rs1.EOF then
'            response.Write(0)
'        end if
'        rs.MoveNext
'     Wend 

'    rs.Close
%>
<html>
    <head>
    <meta charset="utf-8"><meta name="description"><meta name="keywords"><meta name="robots">
    <link rel="stylesheet" href="/css/style.css" type="text/css">
    </head>
    <script type="text/javascript" language="javascript">
        function OnAccept(no)
        {
            var url = "pay_processing.asp?type=1&no=" + no;
            document.location = url;
        }
        function OnWait(no)
        {
            var url = "pay_processing.asp?type=2&no=" + no;
            document.location = url;
        }
        function OnCancel(no)
        {
            var url = "pay_processing.asp?type=3&no=" + no;
            document.location = url;
        }
        function OnDelete(no)
        {
            var url = "delete_no.asp?no=" + no;
            document.location = url;
        }
    </script>
<title>Manage_charge</title>
<body>
    <div class="testing">
        <section class="user">
	        <div class="profile-img">
		        <p><font color="red"><b>Manage</b></font></p>
	        </div>
	        <div class="infos">
	            <table width="300px" cellpadding="0" cellspacing="0">
                <tbody><tr>
                    <td width="100px" class="left">
                        <a style="color:Yellow;">充值申请 : <span id="spChargeRequest_pip"><%=recharge_no%></span></a>
                    </td>
                </tr>
                <tr class="odd">
                    <td width="100px" class="left">
                        <a style="color:Yellow;">充值等待 : <span id="spChargeStandby_pip"><%=recharge_wait%></span></a>
                    </td>
                </tr>
                <tr>
                    <td width="100px" class="left">
                        <a style="color:Yellow;">充值完毕 : <span id="spChargeComplete_pip"><%=recharge_cancel+recharge_accept%></span></a>
                    </td>
                </tr>
	            </tbody></table>
	        </div>
	        <div class="buttons">
	            <span class="button blue"><a href="Index.asp">网页</a></span>
	        </div>
        </section>
    </div>
    <nav style="height: 1079px;">
	    <ul>
		    <li id="liMoneyMng" class="section">
		        <a href="javascript:;"><span class="icon">📜</span> 充提管理</a>
		    </li>
	    </ul>
    </nav>
<section class="content">
	<section class="widget">
		<header>
			<span class="icon">🌄</span>
			<hgroup>
				<h1>充值管理</h1>
			</hgroup>
		</header>
		<div class="content">
		    <table width="100%" border="0" cellpadding="0" cellspacing="0">
		    <tbody>
<!--
    <tr class="odd">
		        <td class="clssearch">
		            <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tbody><tr>
                        <th width="10%" class="left">申请日期</th>
                        <td width="55%" class="left">
                            <input name="ctl00$ManageBodyContent$tbxStartDate" type="text" id="ctl00_ManageBodyContent_tbxStartDate" class="clsinput" onblur="checkDateTime(this, true);" style="width:80px;">
                            &nbsp;~&nbsp;
                            <input name="ctl00$ManageBodyContent$tbxEndDate" type="text" id="ctl00_ManageBodyContent_tbxEndDate" class="clsinput" onblur="checkDateTime(this, true);" style="width:80px;">
                            &nbsp;
                            归属&nbsp;
                            <select name="ctl00$ManageBodyContent$ddlSite" id="ctl00_ManageBodyContent_ddlSite">
	<option selected="selected" value="0">全部</option>
	<option value="1">1</option>
	<option value="2">2</option>

</select>&nbsp;&nbsp;
                        </td>
                        <td width="25%" class="left">
                            <table id="ctl00_ManageBodyContent_rblStatus" class="clscheckbox" border="0">
	<tbody><tr class="odd">
		<td><input id="ctl00_ManageBodyContent_rblStatus_0" type="radio" name="ctl00$ManageBodyContent$rblStatus" value="All" checked="checked">
            <label for="ctl00_ManageBodyContent_rblStatus_0">全部</label></td><td><input id="ctl00_ManageBodyContent_rblStatus_1" type="radio" name="ctl00$ManageBodyContent$rblStatus" value="0">
            <label for="ctl00_ManageBodyContent_rblStatus_1">申请</label></td><td><input id="ctl00_ManageBodyContent_rblStatus_2" type="radio" name="ctl00$ManageBodyContent$rblStatus" value="1">
            <label for="ctl00_ManageBodyContent_rblStatus_2">完毕</label></td><td><input id="ctl00_ManageBodyContent_rblStatus_3" type="radio" name="ctl00$ManageBodyContent$rblStatus" value="2">
            <label for="ctl00_ManageBodyContent_rblStatus_3">取消</label></td><td><input id="ctl00_ManageBodyContent_rblStatus_4" type="radio" name="ctl00$ManageBodyContent$rblStatus" value="3">
            <label for="ctl00_ManageBodyContent_rblStatus_4">等待</label></td>
	</tr>
</tbody></table>
                        </td>
                    </tr>
                    <tr><td colspan="4" class="clssmallspace"></td></tr>
                    <tr class="odd">
                        <th width="10%" class="left">搜索</th>
                        <td class="left">
                            Account&nbsp;<input name="ctl00$ManageBodyContent$tbxLoginID" type="text" id="ctl00_ManageBodyContent_tbxLoginID" class="clsinput">&nbsp;&nbsp;
                            署名&nbsp;<input name="ctl00$ManageBodyContent$tbxNickName" type="text" id="ctl00_ManageBodyContent_tbxNickName" class="clsinput">&nbsp;&nbsp;
                            充值人姓名&nbsp;<input name="ctl00$ManageBodyContent$tbxChargeName" type="text" id="ctl00_ManageBodyContent_tbxChargeName" class="clsinput">&nbsp;&nbsp;
                        </td>
                        <th width="10%" class="left"></th>
                        <td width="40%" class="left">
                            <input type="submit" name="ctl00$ManageBodyContent$btnSearch" value="搜索" id="ctl00_ManageBodyContent_btnSearch" class="button clsbutton green">
                        </td>
                    </tr>
                    </tbody>
    -->
    </table>
		        </td>
		    </tr>
		    <tr><td class="clsspace"></td></tr>
            <tr class="odd">
                <td class="left">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tbody><tr>
<!--
                        <td width="30%" class="left">
                            <input type="button" name="ctl00$ManageBodyContent$btnApply" value="充值处理" onclick="return confirmCheck(MSG_CONFIRMAPPLY);" id="ctl00_ManageBodyContent_btnApply" class="button orange" style="width:80px;">
                            <input type="button" name="ctl00$ManageBodyContent$btnStandby" value="等待处理" onclick="return confirmCheck(MSG_CONFIRMAPPLY);" id="ctl00_ManageBodyContent_btnStandby" class="button blue" style="width:80px;">
                            <input type="button" name="ctl00$ManageBodyContent$btnCancel" value="取消" onclick="return confirmCheck(MSG_CONFIRMCANCEL);" id="ctl00_ManageBodyContent_btnCancel" class="button" style="width:80px;">
                        </td>
-->
                        <td width="40%" class="center">                
                            <b>(10 件&nbsp;/&nbsp;<font color="red"><b>1.00</b>元</font>)</b>
                        </td>
<!--
                        <td width="30%" class="right">
                            <input type="submit" name="ctl00$ManageBodyContent$btnDelete" value="删除" onclick="return confirmCheck(MSG_CONFIRMDELETE);" id="ctl00_ManageBodyContent_btnDelete" class="button" style="width:80px;">
                        </td>
-->
                    </tr>
                    </tbody></table>
                </td>
            </tr>
            <tr class="odd"><td class="clsspace"></td></tr>
            <tr>
                <td>
                    <div>
	<table cellspacing="0" border="0" id="ctl00_ManageBodyContent_gvContent" style="width:100%;border-collapse:collapse;">
		<tbody>
        <tr class="odd">
<!--
			<th class="clstableheader withborder" scope="col"><input type="checkbox" class="clscheckbox" value="All" onclick="checkAll(this)"></th>
-->
            <th class="clstableheader withborder" scope="col">编号</th><th class="clstableheader withborder" scope="col">UserAccount</th>
            <th class="clstableheader withborder" scope="col">充值方式</th><th class="clstableheader withborder" scope="col">账号</th>
            <th class="clstableheader withborder" scope="col">金额</th><th class="clstableheader withborder" scope="col">type</th>
            <th class="clstableheader withborder" scope="col">申请日期</th>
            <th class="clstableheader withborder" scope="col">状态</th><th class="clstableheader withborder" scope="col">处理</th>
		</tr>
<%
    Set rs = Conn.Execute("select * from QPTreasureDB.dbo.ChargeManage where is_delete=0 order by req_datetime desc")
    Dim i
    i = 1
    While Not rs.EOF
%>
        <tr>
<!--
			<td class="clstablecontent withborder" style="width:30px;">
                <input type="checkbox" class="clscheckbox" style="height: auto;" name="chkNo" value="112">
            </td>
-->
            <td class="clstablecontent withborder" style="width:30px;">
                <%=i %>
            </td>
            <td class="clstablecontent withborder" style="width:80px;">
               <a><%=rs.Fields("useraccount") %></a>
            </td>
            <td class="clstablecontent withborder center" style="width:100px;">
                    <font color="green">手动充值</font>
            </td>

                <%
                    if rs.Fields("pay_type") = 1 then
                %>
            <td class="clstablecontent withborder left" style="width:240px;">
                    工商银行
            </td>
                <%        
                    elseif rs.Fields("pay_type") = 2 then
                %>
            <td class="clstablecontent withborder left" style="width:240px;">
                腾讯财付通
            </td>
                <%        
                    elseif rs.Fields("pay_type") = 3 then
                %>
            <td class="clstablecontent withborder left" style="width:240px;">
                支付宝
            </td>
                <%
                    else
                %>
            <td class="clstablecontent withborder left" style="width:240px;">
                兑换
            </td>
                <%
                    end if
                %>
            <td class="clstablecontent withborder left" style="width:90px;">
                    <input type="text" class="clsinput" name="txtReqMoney112" value="<%=rs.Fields("money_amount") %>" readonly>
            </td>
                <%
                    if rs.Fields("recharge_field") = 1 then
                %>
            <td class="clstablecontent withborder" style="width:50px;">
                    <font color="red">充值</font>
            </td>
                <%        
                    elseif rs.Fields("discharge_field") = 1 then
                %>
            <td class="clstablecontent withborder" style="width:50px;">
                    <font color="red">兑换</font>
            </td>
                <%
                    end if
                %>
            <td class="clstablecontent withborder" style="width:100px;"><%=rs.Fields("req_datetime") %></td>
                <%
                    if rs.Fields("accept") = 0 then
                %>
            <td class="clstablecontent withborder" style="width:50px;">
                    <font color="red">申请</font>
            </td>
                <%
                    elseif rs.Fields("accept") = 1 then
                %>
            <td class="clstablecontent withborder" style="width:50px;">
                    <font color="blue">充值</font>
            </td>
                <%
                    elseif rs.Fields("accept") = 2 then
                %>
            <td class="clstablecontent withborder" style="width:50px;">
                    <font color="black">等待</font>
            </td>
                <%
                    elseif rs.Fields("accept") = 3 then
                    %>
            <td class="clstablecontent withborder" style="width:50px;">
                    <font color="black">取消</font>
            </td>
            <%
                    end if
                %>
            <td class="clstablecontent withborder" style="width:130px;">
                    <a onclick="OnAccept(<%=rs.Fields("no")%>);" id="ctl00_ManageBodyContent_gvContent_ctl02_lnkApply" href="javascript:__doPostBack('ctl00$ManageBodyContent$gvContent$ctl02$lnkApply','')">[充值]</a>
                    <a onclick="OnWait(<%=rs.Fields("no")%>);" id="ctl00_ManageBodyContent_gvContent_ctl02_lnkStandby" href="javascript:__doPostBack('ctl00$ManageBodyContent$gvContent$ctl02$lnkStandby','')">[等待]</a>
                    <a onclick="OnCancel(<%=rs.Fields("no")%>);" id="ctl00_ManageBodyContent_gvContent_ctl02_lnkCancel" href="javascript:__doPostBack('ctl00$ManageBodyContent$gvContent$ctl02$lnkCancel','')">[取消]</a>
                    <a onclick="OnDelete(<%=rs.Fields("no")%>);" id="ctl00_ManageBodyContent_gvContent_ctl02_lnkCancel" href="javascript:__doPostBack('ctl00$ManageBodyContent$gvContent$ctl02$lnkCancel','')">[删除]</a> 
            </td>
		</tr>
<%
    i = i + 1
    rs.MoveNext
    Wend 
%>
	</tbody></table>
            </div>
            </td>
            </tr>
		    </tbody></table>
		</div>
	</section>
</section>
    <div id="divMenu" class="clsmenu" style="display: none" onmouseover="menuShowAction()" onmouseout="menuHideAction()"></div>
    <div id="divPlaySound"><embed name="objPlaySoundName" src="/images/playsound.swf" quality="high" style="width:1px; height:1px;" align="middle" allowscriptaccess="always" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?p1_prod_version=shockwaveflash" autoplay="false" loop="true"></div>
<%
    Set rs = Nothing
    Conn.Close
    Set Conn = Nothing
%>
</body>
</html>