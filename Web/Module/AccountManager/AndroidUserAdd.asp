<!--#include file="../../CommonFun.asp"-->
<!--#include file="../../function.asp"-->
<!--#include file="../../GameConn.asp"-->
<!--#include file="../../Cls_Page.asp"-->
<!--#include file="../../conn.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
    <title></title>
    <link href="../../Css/layout.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../../Js/common.js"></script>
    <script type="text/javascript" src="../../Js/comm.js"></script>
    <script type="text/javascript" src="../../Js/Check.js"></script>
    <script type="text/javascript" src="../../Js/Calendar.js"></script>
    <script type="text/javascript" src="../../Js/Sort.js"></script>
    <script type="text/javascript" src="../../Js/My97DatePicker/WdatePicker.js"></script>
    <script type="text/javascript">
    function setAction(type)
    {
        if(type==1)
        {
            document.getElementById("form2").action="?action=save";
            document.getElementById("oneAdd").style.display="";
            document.getElementById("moreAdd").style.display="none";
        }
        if(type==2)
        {
            document.getElementById("form2").action="?action=savemore";
            document.getElementById("oneAdd").style.display="none";
            document.getElementById("moreAdd").style.display="";
        }
    }
    </script>
</head>
<body>
   <!-- 头部菜单 Start -->
    <table width="100%" border="0" cellpadding="0" cellspacing="0" class="title">
        <tr>
            <td width="19" height="25" valign="top"  class="Lpd10"><div class="arr"></div></td>
            <td width="1232" height="25" valign="top" align="left">目前操作功能：增加机器人</td>
        </tr>
    </table>
    <!-- 头部菜单 End -->
    <% 
        Select case lcase(Request("action"))
            case "save"
            Call Save()
             case "savemore"
            Call SaveMore()
            
        End Select
        
        Sub Save()
            Dim dbBase,rs,sql,userName,userID
            Dim MinPlayDraw,MaxPlayDraw,MinTakeScore,MaxTakeScore,MinReposeTime,MaxReposeTime,AndroidNote,Nullity,CreateDate 
            Dim ServiceGenderArr,ServiceGender,i
            Dim ServiceTimeArr,ServiceTime,j
            MinPlayDraw = CxGame.GetInfo(1,"MinPlayDraw")
            MaxPlayDraw = CxGame.GetInfo(1,"MaxPlayDraw")
            MinTakeScore = CxGame.GetInfo(0,"MinTakeScore")
            If MinTakeScore="" Then
                MinTakeScore=0
            End If 
            MaxTakeScore = CxGame.GetInfo(0,"MaxTakeScore")
            If MaxTakeScore="" Then
                MaxTakeScore=0
            End If
            MinReposeTime = CxGame.GetInfo(1,"MinReposeTime")
            MaxReposeTime = CxGame.GetInfo(1,"MaxReposeTime")    
            '机器人类型        
            ServiceGenderArr = Split(Request("ServiceGender"),",")
            ServiceGender=0
            For i=0 To UBound(ServiceGenderArr)
                ServiceGender = ServiceGender Or ServiceGenderArr(i)
            Next
            '服务时间
            ServiceTimeArr = Split(Request("ServiceTime"),",")
            ServiceTime=0
            For j=0 To UBound(ServiceTimeArr)
                ServiceTime = ServiceTime Or ServiceTimeArr(j)
            Next
            
            AndroidNote = CxGame.GetInfo(0,"AndroidNote")
            Nullity = CxGame.GetInfo(1,"Nullity")                               
            dbBase = Split(Request("RoomID"),",")
            userName=Request("Accounts")
            Call ConnectGame(dbBase(2))
            Set rs=Server.CreateObject("Adodb.RecordSet")
            If userName<>"" Then
                sql = "select userid from QPAccountsDBLink.QPAccountsDB.dbo.AccountsInfo where Accounts='"&userName&"'"
                rs.Open sql,GameConn,1,3
                If rs.RecordCount=0 Then
                    CxGame.MessageBox("该机器人不存在！")
                Else
                    userID=rs("UserID")
                    sql = "select * from AndroidManager where userid =(select userid from QPAccountsDBLink.QPAccountsDB.dbo.AccountsInfo where Accounts='"&userName&"')"
                    Set rs=Server.CreateObject("Adodb.RecordSet")
                    rs.Open sql,GameConn,1,3
                    If rs.RecordCount>0 Then
                        CxGame.MessageBox("该机器人已经被添加！")
                    Else
                    Call ConnectGame(dbBase(2))
                        sql = "insert into AndroidManager(UserID,ServerID,MinPlayDraw,MaxPlayDraw,MinTakeScore,MaxTakeScore,MinReposeTime,MaxReposeTime,ServiceTime,ServiceGender,AndroidNote,Nullity)" &_
			                "values("&userID &"," & dbBase(1) & ","&MinPlayDraw&","&MaxPlayDraw&","&MinTakeScore&","&MaxTakeScore&","&MinReposeTime&","&MaxReposeTime&","&ServiceTime&","&ServiceGender&",'"&AndroidNote&"',"&Nullity&")"
			        GameConn.execute sql
			        Call CxGame.MessageBoxReload("添加机器人"&userName&"成功！","AndroidUserInfo.asp")
                    End If
                End If
            Else
                CxGame.MessageBox("请输入机器名！")
            End If
            Set rs=Nothing
            Call CloseGame()
        End Sub
        
         Sub SaveMore()
            Dim dbBase,rs,sql,counts,userID,dbBaseLink,i,j
            j=0
            Dim MinPlayDraw,MaxPlayDraw,MinTakeScore,MaxTakeScore,MinReposeTime,MaxReposeTime,ServiceGender,AndroidNote,Nullity,CreateDate 
            Dim ServiceGenderArr2,ServiceGender2,RL_i
            Dim ServiceTimeArr2,ServiceTime2,RL_j
            MinPlayDraw = CxGame.GetInfo(1,"MinPlayDraw")
            MaxPlayDraw = CxGame.GetInfo(1,"MaxPlayDraw")
            MinTakeScore = CxGame.GetInfo(0,"MinTakeScore")
            If MinTakeScore="" Then
                MinTakeScore=0
            End If 
            MaxTakeScore = CxGame.GetInfo(0,"MaxTakeScore")
            If MaxTakeScore="" Then
                MaxTakeScore=0
            End If
            MinReposeTime = CxGame.GetInfo(1,"MinReposeTime")
            MaxReposeTime = CxGame.GetInfo(1,"MaxReposeTime")
            '机器人类型        
            ServiceGenderArr2 = Split(Request("ServiceGender"),",")
            ServiceGender2=0
            For RL_i=0 To UBound(ServiceGenderArr2)
                ServiceGender2 = ServiceGender2 Or ServiceGenderArr2(RL_i)
            Next
            '服务时间
            ServiceTimeArr2 = Split(Request("ServiceTime"),",")
            ServiceTime2=0
            For RL_j=0 To UBound(ServiceTimeArr2)
                ServiceTime2 = ServiceTime2 Or ServiceTimeArr2(RL_j)
            Next
            
            AndroidNote = CxGame.GetInfo(0,"AndroidNote")
            Nullity = CxGame.GetInfo(1,"Nullity")        
            
            dbBase = Split(Request("RoomID2"),",")
            counts= Request("counts")
            If counts<>"" Then
                Call ConnectGame(dbBase(2))
                Set rs=Server.CreateObject("Adodb.RecordSet")
                sql = "SELECT top "&counts&" UserID from QPAccountsDBLink.QPAccountsDB.dbo.AccountsInfo WHERE UserID not in "
                sql=sql&"(SELECT b.UserID from AndroidManager a,QPAccountsDBLink.QPAccountsDB.dbo.AccountsInfo b where a.UserID=b.UserID) and IsAndroid=1 order by newid()"
                rs.Open sql,GameConn,1,3
                Call CloseGame()
                
                If rs.RecordCount=0 Then                     
                    CxGame.MessageBox("无可用机器人！")
                Else
                    Call ConnectGame(dbBase(2))
                    do until rs.eof 
                    Set userID=rs("UserID")
                    sql = "insert into AndroidManager(UserID,ServerID,MinPlayDraw,MaxPlayDraw,MinTakeScore,MaxTakeScore,MinReposeTime,MaxReposeTime,ServiceTime,ServiceGender,AndroidNote,Nullity)" &_
		                "values("&userID &"," & dbBase(1) & ","&MinPlayDraw&","&MaxPlayDraw&","&MinTakeScore&","&MaxTakeScore&","&MinReposeTime&","&MaxReposeTime&","&ServiceTime2&","&ServiceGender2&",'"&AndroidNote&"',"&Nullity&")"
	                GameConn.execute sql
                    rs.movenext
                    j=j+1
                    loop
                    Call CxGame.MessageBoxReload("成功增加"&j&"个机器人！","AndroidUserInfo.asp")
                    Set rs=Nothing
                    Call CloseGame()
                End If
             End If
         End Sub
    %>
    <form name="form2" id="form2" method="post" action="?action=save">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td class="titleOpBg Lpd10">
                    <input type="button" class="btn wd1" onclick="Redirect('AndroidUserInfo.asp')" value="返回" />   
                    <input type="submit" value="保存" class="btn wd1" onclick="return CheckCounts();"/>   
                   
                </td>
            </tr>
        </table>
        <table  width="100%" class="listBg2" cellpadding="0" cellspacing="0">
            <tr>
                 <td class="listTdLeft">
                    新增方式：
                </td>
                <td>
                    <input type="radio" name="rb" id="rb1" value="单笔新增" checked="checked" onclick="setAction('1')" /><label for="rb1">单笔新增</label>
                    <input type="radio" name="rb" id="rb2" value="批量新增" onclick="setAction('2')"/><label for="rb2">批量新增</label>
                </td>
            </tr>
        </table>
        <table id="oneAdd" width="100%" class="listBg2" cellpadding="0" cellspacing="0" >
            <tr>
                <td class="listTdLeft">
                    请填入机器人名：
                </td>
                <td>
                    <input type="text" name="Accounts" id="Accounts" class="text" style="width: 150px" />
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">
                    请选择游戏房间：
                </td>
                <td>
                     <select name="RoomID" id="Select1" style="width:157px;">
                     <%
                        Call ConnectGame("QPPlatformDB")
                          Dim rs,sql,GameName
                                Set rs=Server.CreateObject("Adodb.RecordSet")
                                sql = "select * from GameRoomInfo(nolock) ORDER BY SortID"
                                rs.Open sql,GameConn,1,3
            			 
		  	            IF Not rs.Eof Then
				            Do While Not rs.Eof
					            GameName=GetGameNameByID(rs("GameID"))
            				
		            %>
                    <option value="<%= rs("GameID") %>,<%= rs("ServerID")%>,<%= rs("DataBaseName")%>"><%=rs("ServerName") %></option>
                    <%
					    rs.MoveNext
				            Loop
			            Else
			
		            %>
                    <option value="0">没有任何房间</option>
                    <%
                        Set rs = Nothing
                        Call CloseGame()
			        End IF
		
		            %>
                  </select>
                   
                </td>
            </tr>
        </table>
         <table id="moreAdd"  width="100%" style="display:none;" class="listBg2" cellpadding="0" cellspacing="0">
             <tr>
                <td class="listTdLeft">
                    请填入新增机器人个数：
                </td>
                <td style="height: 24px;">
                    <input type="text" id="counts" name="counts" class="text"  style="width: 150px" />
                    
                    </td>
                    </tr>
                    <tr>
                    <td class="listTdLeft">
                    请选择游戏房间：
                    </td>
                    <td>
                   <select name="RoomID2" id="Select2" >
                   <% 
                   Call ConnectGame("QPPlatformDB")
                        Set rs=Server.CreateObject("Adodb.RecordSet")
                        sql = "select * from GameRoomInfo(nolock) ORDER BY SortID"
                        rs.Open sql,GameConn,1,3
			 
		  	        IF Not rs.Eof Then
				        Do While Not rs.Eof
					        GameName=GetGameNameByID(rs("GameID"))
				  %>
                  <option value="<%= rs("GameID") %>,<%= rs("ServerID")%>,<%= rs("DataBaseName")%>"><%=rs("ServerName") %></option>
                  <%
					rs.MoveNext
				        Loop
			        Else
		          %>
                  <option value="0">没有任何房间</option>
                  <%
                    Set rs = Nothing
                    Call CloseGame()
			        End IF
		
		          %>
		          </select>
                      
                </td>
            </tr>
        </table>
        <table id="Table1"  width="100%" class="listBg2" cellpadding="0" cellspacing="0">
            <tr>
                <td class="listTdLeft">最少局数：</td>
                <td>
                    <input name="MinPlayDraw" type="text" class="text" />    
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">最大局数：</td>
                <td>
                    <input name="MaxPlayDraw" type="text" class="text" />           
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">最少分数：</td>
                <td>
                    <input name="MinTakeScore" type="text" class="text" />              
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">最高分数：</td>
                <td>
                    <input name="MaxTakeScore" type="text" class="text" />              
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">最小游戏时间：</td>
                <td>
                    <input name="MinReposeTime" type="text" class="text" />              
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">最大游戏时间：</td>
                <td>
                    <input name="MaxReposeTime" type="text" class="text" />              
                </td>
            </tr>            
            <tr>
                <td class="listTdLeft">服务类型：</td>
                <td> 
                    <input name="ServiceGender" id="chkServiceGender1" type="checkbox" value="1" /><label for="chkServiceGender1">相互模拟</label> 
                    <input name="ServiceGender" id="chkServiceGender2" type="checkbox" value="2" /><label for="chkServiceGender2">被动陪打</label>            
                    <input name="ServiceGender" id="chkServiceGender3" type="checkbox" value="4" /><label for="chkServiceGender3">主动陪打</label>      
                    <span class="lan">(不勾选为挂线机器人【机器人直接挂在房间，不进行任何操作。】)</span>  
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">服务时间：<br />（全选<input type="checkbox"  onclick="SelectAllTable(this.checked,'serviceTime');"/>）</td>
                <td> 
                    <table border="0" style="padding:5px 5px 5px 0;" cellpadding="0" cellspacing="0" id="serviceTime">
                    <tr>
                        <td><input name="ServiceTime" id="chkServiceTime1" type="checkbox" value="1" /><label for="chkServiceTime1">0:00-1:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime2" type="checkbox" value="2" /><label for="chkServiceTime2">1:00-2:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime3" type="checkbox" value="4" /><label for="chkServiceTime3">2:00-3:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime4" type="checkbox" value="8" /><label for="chkServiceTime4">3:00-4:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime5" type="checkbox" value="16" /><label for="chkServiceTime5">4:00-5:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime6" type="checkbox" value="32" /><label for="chkServiceTime6">5:00-6:00</label></td>
                    </tr>
                    <tr>
                        <td><input name="ServiceTime" id="chkServiceTime7" type="checkbox" value="64" /><label for="chkServiceTime7">6:00-7:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime8" type="checkbox" value="128" /><label for="chkServiceTime8">7:00-8:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime9" type="checkbox" value="256" /><label for="chkServiceTime9">8:00-9:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime10" type="checkbox" value="512" /><label for="chkServiceTime10">9:00-10:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime11" type="checkbox" value="1024" /><label for="chkServiceTime11">10:00-11:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime12" type="checkbox" value="2048" /><label for="chkServiceTime12">11:00-12:00</label></td>
                    </tr>
                    <tr>
                        <td><input name="ServiceTime" id="chkServiceTime13" type="checkbox" value="4096" /><label for="chkServiceTime13">12:00-13:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime14" type="checkbox" value="8192" /><label for="chkServiceTime14">13:00-14:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime15" type="checkbox" value="16384" /><label for="chkServiceTime15">14:00-15:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime16" type="checkbox" value="32768" /><label for="chkServiceTime16">15:00-16:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime17" type="checkbox" value="65536" /><label for="chkServiceTime17">16:00-17:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime18" type="checkbox" value="131072" /><label for="chkServiceTime18">17:00-18:00</label></td>
                    </tr>
                    <tr>
                        <td><input name="ServiceTime" id="chkServiceTime19" type="checkbox" value="262144" /><label for="chkServiceTime19">18:00-19:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime20" type="checkbox" value="524288" /><label for="chkServiceTime20">19:00-20:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime21" type="checkbox" value="1048576" /><label for="chkServiceTime21">20:00-21:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime22" type="checkbox" value="2097152" /><label for="chkServiceTime22">21:00-22:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime23" type="checkbox" value="4194304" /><label for="chkServiceTime23">22:00-23:00</label></td>
                        <td><input name="ServiceTime" id="chkServiceTime24" type="checkbox" value="8388608" /><label for="chkServiceTime24">23:00-24:00</label></td>
                    </tr>
                    </table>  
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">备注信息：</td>
                <td>
                    <input name="AndroidNote" type="text" class="text" style="width:300px;" />              
                </td>
            </tr>
            <tr>
                <td class="listTdLeft">禁用状态：</td>
                <td>
                    <input name="in_Nullity" id="Checkbox1" type="radio" value="0" checked="checked" /><label for="Checkbox1">启用</label> 
                    <input name="in_Nullity" id="Checkbox2" type="radio" value="1" /><label for="Checkbox2">冻结</label> 
                </td>
            </tr>
        </table>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td class="titleOpBg Lpd10">
                    <input type="button" class="btn wd1" onclick="Redirect('AndroidUserInfo.asp')" value="返回" />   
                    <input type="submit" value="保存" class="btn wd1" onclick="return CheckCounts();" />  
                 
                </td>
            </tr>
        </table>
    </form>
   <%
    Function GetGameNameByID( KindID)
     Call ConnectGame("QPPlatformDB")
    Dim lRs, lSql, GameName
    Set lRs =Server.CreateObject("Adodb.RecordSet")
    lSql="select KindName from GameKindItem where KindID = " & KindID
    lRs.Open lSql,GameConn,1,3
    If Not lRs.Eof Then
	    IF isnull(lRs(0)) Then
		    GetGameNameByID = ""
	    Else
		    GetGameNameByID = lRs(0)
	    End IF
    Else
	    GetGameNameByID = ""
    End IF

    lRs.Close
    Set lRs = Nothing
    End Function
   
    %>
</body>
</html>
<script type="text/javascript">
    function CheckCounts()
    {
        var isType=IsPositiveInt( document.getElementById("counts").value);
        if(document.getElementById("rb2").checked)
        {
            if(!isType)
            {
                alert("请输入整数！");
                return false
            }
            return true;
        }
    }
</script>