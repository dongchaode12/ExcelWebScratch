Attribute VB_Name = "ģ��1"
Public Function GetInfo(StrMobile As String) As String
    '��������
    Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
    
    'xmlHttp.open "����ʽ","��ַ",flase
    '��������
    xmlHttp.Open "GET", "https://sp0.baidu.com/8aQDcjqpAAV3otqbppnN2DJv/api.php?resource_name=guishudi&query=" & StrMobile, flase
    xmlHttp.send

      
    '�ȴ���Ӧ
    Do While xmlHttp.ReadyState <> 4
       DoEvents
    Loop
    
    
    '������Ӧ
    Dim strReturn, see As String
    strReturn = xmlHttp.responsetext
    
    'see = "https://tcc.taobao.com/cc/json/mobile_tel_segment.htm?tel=" & StrMobile  ���һ������������鷢�͵���ַ�Ƿ���ȷ������
    
    '��������
    
    Dim strCity, strPro, strCom As String
    strCity = Replace(Split(Split(strReturn, ",")(7), ":")(1), """", "")   '����split�����ָ����ݣ�����replace�����滻""����
    strPro = Replace(Split(Split(strReturn, ",")(9), ":")(1), """", "")
    strCom = Replace(Split(Split(strReturn, ",")(12), ":")(1), """", "")
    
    GetInfo = strCom & "-" & IIf(strPro = strCity, strCity, strPro & strCity) '����iif������ж�ֱϽ�е������Ƿ���ͬ
    
    
    'GetInfo = strReturn
    
    
    
    
    
    
End Function
