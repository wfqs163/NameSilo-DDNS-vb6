VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "namesilo_ddns"
   ClientHeight    =   855
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   855
   ScaleWidth      =   3600
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Update DNS record"
      Height          =   540
      Left            =   885
      TabIndex        =   0
      Top             =   180
      Width           =   1920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mc_strIniFileName As String = "config.ini"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Dim APIkey As String
Dim Domain As String
Dim RRHOST As String
Dim SILENT As Integer
Dim LocalIP As String
Dim RemoteIP As String
Dim DNSRecord_ID As String

Private Sub Form_load()
    APIkey = GetIni("namesilo_ddns", "apikey")
    Domain = GetIni("namesilo_ddns", "domain")
    RRHOST = GetIni("namesilo_ddns", "rrhost")
    SILENT = GetIni("namesilo_ddns", "silent")
    
    Dim HttpRequest As New WinHttp.WinHttpRequest
    Dim HttpResponse As String
    
    HttpRequest.Open "GET", "https://www.namesilo.com/api/dnsListRecords?version=1&type=xml&key=" & APIkey & "&domain=" & Domain, False
    HttpRequest.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
    HttpRequest.Send
    HttpResponse = HttpRequest.ResponseText
    
    Dim DomXML As New MSXML2.DOMDocument
    Dim DomNode As IXMLDOMNode
    Dim resource As IXMLDOMElement
    
    DomXML.loadXML HttpResponse
    Set DomNode = DomXML.selectNodes("namesilo").Item(0)
    LocalIP = DomNode.selectSingleNode("request").selectSingleNode("ip").Text
    
    For Each resource In DomNode.selectSingleNode("reply").selectNodes("resource_record")
        If resource.selectSingleNode("host").Text = RRHOST & "." & Domain Then
            RemoteIP = resource.selectSingleNode("value").Text
            DNSRecord_ID = resource.selectSingleNode("record_id").Text
        End If
    Next
    
    HttpRequest.Open "GET", "https://www.namesilo.com/api/dnsUpdateRecord?version=1&type=xml&key=" & APIkey & "&domain=" & Domain & "&rrid=" & DNSRecord_ID & "&rrhost=" & RRHOST & "&rrvalue=" & LocalIP & "&rrttl=3600", False
    HttpRequest.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
    HttpRequest.Send
    HttpResponse = HttpRequest.ResponseText
    If DomXML.selectNodes("namesilo").Item(0).selectSingleNode("reply").selectSingleNode("code").Text = "300" Then
        If SILENT = 0 Then MsgBox "Update DNS RECORD successed!"
    Else
        MsgBox "Update DNS RECORD Failed!" & vbCrLf & HttpResponse
    End If
    
    End
End Sub





Public Function GetIni(appName As String, keyName As String) As String

    Dim strDefault As String
    Dim lngBuffLen As Long
    Dim strResu As String
    Dim X As Long
    Dim strIniFile As String
      
    If Right(App.Path, 1) = "\" Then
        strIniFile = App.Path & mc_strIniFileName
    Else
        strIniFile = App.Path & "\" & mc_strIniFileName
    End If
      
    strResu = String(1025, vbNullChar): lngBuffLen = 1025
    strDefault = ""
    X = GetPrivateProfileString(appName, keyName, strDefault, strResu, lngBuffLen, strIniFile)
    Debug.Print X
    Debug.Print strResu
    GetIni = Left(Trim(strResu), X)
      
End Function

