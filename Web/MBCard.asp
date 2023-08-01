<%
'默认固定模版
dim A1,A2,A3,A4,B1,B2,B3,B4,C1,C2,C3,C4,D1,D2,D3,D4,E1,E2,E3,E4,F1,F2,F3,F4
A1=9
A2=15
A3=24
A4=268
B1=1
B2=88
B3=25
B4=76
C1=185
C2=28
C3=65
C4=205
D1=72
D2=39
D3=11
D4=31
E1=52
E2=220
E3=7
E4=105
F1=99
F2=117
F3=93
F4=48

Dim a,b,c,d,e,f,arra(23)
arra(0) = "A1"
arra(1) = "A2"
arra(2) = "A3"
arra(3) = "A4"
arra(4) = "B1"
arra(5) = "B2"
arra(6) = "B3"
arra(7) = "B4"
arra(8) = "C1"
arra(9) = "C2"
arra(10)= "C3"
arra(11)= "C4"
arra(12)= "D1"
arra(13)= "D2"
arra(14)= "D3"
arra(15)= "D4"
arra(16)= "E1"
arra(17)= "E2"
arra(18)= "E3"
arra(19)= "E4"
arra(20)= "F1"
arra(21)= "F2"
arra(22)= "F3"
arra(23)= "F4"
Randomize
a = Int(23*Rnd())
Randomize
b = Int(23*Rnd())
Randomize
c = Int(23*Rnd())
Randomize
d = Int(23*Rnd())
Randomize
e = Int(23*Rnd())
Randomize
f = Int(23*Rnd())

Dim AAA,BBB,CCC
AAA=arra(a)
BBB=arra(b)
CCC=arra(c)
%>
 
<% 
    '获取密保信息
    Function GetPasswordNum(objA,objB)
        Dim rValue
        rValue = Int(Int(objA)/Int(objB)) Mod 1000
        If Len(rValue)=1 Then
            rValue=rValue&"00"
        End If
        
        If Len(rValue)=2 Then
            rValue=rValue&"0"
        End If
        GetPasswordNum= rValue
    End Function
    
    '得到新的密保信息
    Function GetNewPassWordID(objP)
        Dim rValue
        rValue = Mid(objP,1,3)&" "&Mid(objP,4,3)&" "&Mid(objP,7,3)
        GetNewPassWordID = rValue
    End Function
    
    Rem 生成指定长度的数字随机数
    Function GetRandomNum(cardLength)
        Dim i,rValue
        for i=1 To cardLength
            rValue = rValue&GetRand(1,9)
        Next
        GetRandomNum = rValue
    End Function
    
    '生成一个随机数
    Function GetRand(min,max)
        Dim rValue
        Randomize
        rValue = Int((max - min + 1) * Rnd + min) 
        GetRand = rValue
    End Function
%>