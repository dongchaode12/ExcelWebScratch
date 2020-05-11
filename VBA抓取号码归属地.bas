Attribute VB_Name = "模块1"
Public Function GetInfo(StrMobile As String) As String
    '创建对象
    Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
    
    'xmlHttp.open "请求方式","网址",flase
    '发送请求
    xmlHttp.Open "GET", "https://sp0.baidu.com/8aQDcjqpAAV3otqbppnN2DJv/api.php?resource_name=guishudi&query=" & StrMobile, flase
    xmlHttp.send

      
    '等待响应
    Do While xmlHttp.ReadyState <> 4
       DoEvents
    Loop
    
    
    '接收响应
    Dim strReturn, see As String
    strReturn = xmlHttp.responsetext
    
    'see = "https://tcc.taobao.com/cc/json/mobile_tel_segment.htm?tel=" & StrMobile  添加一个变量用来检查发送的网址是否正确？？？
    
    '处理数据
    
    Dim strCity, strPro, strCom As String
    strCity = Replace(Split(Split(strReturn, ",")(7), ":")(1), """", "")   '利用split函数分割数据，利用replace函数替换""符号
    strPro = Replace(Split(Split(strReturn, ",")(9), ":")(1), """", "")
    strCom = Replace(Split(Split(strReturn, ",")(12), ":")(1), """", "")
    
    GetInfo = strCom & "-" & IIf(strPro = strCity, strCity, strPro & strCity) '利用iif语句来判断直辖市的名字是否相同
    
    
    'GetInfo = strReturn
    
    
    
    
    
    
End Function
